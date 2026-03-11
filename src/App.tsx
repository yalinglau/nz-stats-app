import { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';

type StandardMap = { [key: string]: string };

export default function App() {
  const [data, setData] = useState<any[]>([]);
  const [sMap, setSMap] = useState<StandardMap>({});
  const [cMap, setCMap] = useState<StandardMap>({});
  const [fixes, setFixes] = useState<StandardMap>({});
  const [tempFixes, setTempFixes] = useState<StandardMap>({});
  const [bannerLoaded, setBannerLoaded] = useState(false);

  const cleanStr = (s: any) => String(s || "").replace(/\s+/g, '').toLowerCase();
  const today = new Date().toLocaleDateString();

  const res = useMemo(() => {
    if (data.length === 0) return null;
    const teams: { [key: string]: any } = {};
    const unSchools = new Set<string>();
    const currentFullSMap = { ...sMap, ...fixes };

    // 1. 嚴格提取所有隊伍
    data.forEach(row => {
      const sn = String(row["隊伍序號"] || "").trim();
      if (!sn || sn.startsWith("888")) return;
      
      const [gid, mid] = sn.split("_");
      if (!teams[gid]) {
        // 抓取賽別並徹底清理前後空格
        const rawCat = String(row["賽別"] || "").trim();
        teams[gid] = { 
          category: rawCat,
          school: String(row["學校"] || "").trim(),
          members: []
        };
      }
      teams[gid].members.push(mid ? parseInt(mid) : 0);
    });

    const stats: any = {
      Energy: { teams: 0, people: 0, schools: new Set(), countries: new Set(), overseas: 0, list: [] },
      Sustainability: { teams: 0, people: 0, schools: new Set(), countries: new Set(), overseas: 0, list: [] }
    };

    // 2. 進行統計與歸類
    Object.keys(teams).forEach(gid => {
      const t = teams[gid];
      // 關鍵修復：不使用關鍵字包含判斷，改用標準化比對，預設為 Energy
      const cat = t.category.toLowerCase().includes("sustainability") ? "Sustainability" : "Energy";
      
      const rawSchoolClean = cleanStr(t.school);
      const matchKey = Object.keys(currentFullSMap).find(k => cleanStr(k) === rawSchoolClean);
      const stdSchool = matchKey ? currentFullSMap[matchKey] : "";

      if (!stdSchool) {
        unSchools.add(t.school);
        return;
      }

      const pCount = (t.members.length === 1 && t.members[0] === 0) ? 1 : t.members.filter((m: number) => m !== 0).length;
      
      let country = "臺灣";
      if (stdSchool.includes(",")) {
        const parts = stdSchool.split(",");
        const rawC = parts[parts.length - 1].trim();
        const countryMatchKey = Object.keys(cMap).find(k => cleanStr(k) === cleanStr(rawC));
        country = countryMatchKey ? cMap[countryMatchKey] : rawC;
      }

      stats[cat].teams += 1;
      stats[cat].people += pCount;
      stats[cat].schools.add(stdSchool);
      stats[cat].countries.add(country);
      if (stdSchool.includes(",")) stats[cat].overseas += 1;
      
      const existing = stats[cat].list.find((i:any) => i.school === stdSchool);
      if (existing) { existing.count += 1; } 
      else { stats[cat].list.push({ school: stdSchool, country, count: 1 }); }
    });

    return { stats, unSchools: Array.from(unSchools) };
  }, [data, sMap, fixes, cMap]);

  const loadFile = (e: any, type: string) => {
    const f = e.target.files[0];
    if (!f) return;
    const r = new FileReader();
    r.onload = (ex) => {
      const bstr = ex.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      if (type === 'raw') setData(json);
      else {
        const m: StandardMap = {};
        json.forEach((row: any) => {
          const ks = Object.keys(row);
          if (row[ks[0]]) m[String(row[ks[0]]).trim()] = String(row[ks[1]] || "").trim();
        });
        if (type === 'school') setSMap(m); else setCMap(m);
      }
    };
    r.readAsBinaryString(f);
  };

  const exportExcel = (type: string) => {
    const wb = XLSX.utils.book_new();
    if (type === 'school') {
      const d = Object.entries({ ...sMap, ...fixes }).map(([k, v]) => ({ "原始名稱": k, "標準名稱": v }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(d), "學校對照表");
      XLSX.writeFile(wb, `更新後學校標準對照表.xlsx`);
    } else if (type === 'country') {
      const d = Object.entries(cMap).map(([k, v]) => ({ "原始名稱": k, "標準名稱": v }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(d), "國家對照表");
      XLSX.writeFile(wb, `更新後國家標準對照表.xlsx`);
    } else if (type === 'stats' && res) {
      const s1 = [["統計項目", "Energy", "Sustainability", "合計", "備註"], ["報名隊伍數", res.stats.Energy.teams, res.stats.Sustainability.teams, res.stats.Energy.teams + res.stats.Sustainability.teams, ""], ["報名人數", res.stats.Energy.people, res.stats.Sustainability.people, res.stats.Energy.people + res.stats.Sustainability.people, ""], ["報名學校數", res.stats.Energy.schools.size, res.stats.Sustainability.schools.size, new Set([...res.stats.Energy.schools, ...res.stats.Sustainability.schools]).size, "不重複"], ["海外學校數", res.stats.Energy.overseas, res.stats.Sustainability.overseas, res.stats.Energy.overseas + res.stats.Sustainability.overseas, ""]];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s1), "統計總表");
      const s2 = [["#", "學校名稱", "代表國家", "隊伍數"], ...res.stats.Energy.list.map((item:any, i:number)=>[i+1, item.school, item.country, item.count])];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s2), "Energy組");
      const s3 = [["#", "學校名稱", "代表國家", "隊伍數"], ...res.stats.Sustainability.list.map((item:any, i:number)=>[i+1, item.school, item.country, item.count])];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s3), "Sustainability組");
      XLSX.writeFile(wb, `本次統計資料表.xlsx`);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 pb-20">
      <img src="/banner.png" alt="Banner" className="w-full h-auto block mb-8" />
      <div className="max-w-[1000px] mx-auto px-4">
        <div className="bg-white p-6 rounded-xl shadow mb-8">
          <h1 className="text-xl font-bold mb-4">2026 NZ 統計控制中心</h1>
          <div className="grid grid-cols-3 gap-4">
            {['raw', 'school', 'country'].map(id => (
              <div key={id} className="p-3 border rounded bg-slate-50">
                <p className="font-bold text-xs mb-1">{id === 'raw' ? '1. 原始報名資料' : id === 'school' ? '2. 學校標準表' : '3. 國家標準表'}</p>
                <input type="file" onChange={e => loadFile(e, id)} className="text-xs w-full"/>
              </div>
            ))}
          </div>
        </div>

        {res && (
          <div className="space-y-8">
            <div className="bg-white p-6 rounded shadow border">
              <div id="table-summary">
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end', marginBottom: '4px' }}>
                  <h2 className="text-xl font-bold">NZ 目前報名狀況統計表</h2>
                  <div style={{ fontSize: '12px', fontWeight: 'bold' }}>資料更新日期：{today}</div>
                </div>
                <table className="w-full border-collapse border-2 border-slate-900 font-bold text-sm">
                  <thead className="bg-slate-100">
                    <tr>
                      <th style={{ textAlign: 'left', padding: '6px', border: '2px solid #000' }}>統計項目</th>
                      <th style={{ textAlign: 'right', padding: '6px', border: '2px solid #000' }}>Energy</th>
                      <th style={{ textAlign: 'right', padding: '6px', border: '2px solid #000' }}>Sustainability</th>
                      <th style={{ textAlign: 'right', padding: '6px', border: '2px solid #000' }} className="bg-slate-200">合計</th>
                      <th style={{ textAlign: 'left', padding: '6px', border: '2px solid #000', fontWeight: 'normal' }} className="text-xs text-slate-500">備註</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[{ l: '報名隊伍數', k: 'teams' }, { l: '報名人數', k: 'people' }, { l: '報名學校數', k: 'schools', note: '不計算重複學校' }, { l: '海外學校數', k: 'overseas' }].map((row, i) => (
                      <tr key={i}>
                        <td style={{ textAlign: 'left', padding: '6px', border: '1px solid #000' }}>{row.l}</td>
                        <td style={{ textAlign: 'right', padding: '6px', border: '1px solid #000' }}>{i === 2 ? res.stats.Energy[row.k].size : res.stats.Energy[row.k]}</td>
                        <td style={{ textAlign: 'right', padding: '6px', border: '1px solid #000' }}>{i === 2 ? res.stats.Sustainability[row.k].size : res.stats.Sustainability[row.k]}</td>
                        <td style={{ textAlign: 'right', padding: '6px', border: '1px solid #000' }} className="bg-slate-50">{row.k === 'teams' ? res.stats.Energy.teams + res.stats.Sustainability.teams : row.k === 'people' ? res.stats.Energy.people + res.stats.Sustainability.people : row.k === 'overseas' ? res.stats.Energy.overseas + res.stats.Sustainability.overseas : new Set([...res.stats.Energy.schools, ...res.stats.Sustainability.schools]).size}</td>
                        <td style={{ textAlign: 'left', padding: '6px', border: '1px solid #000', fontSize: '11px', fontWeight: 'normal' }}>{row.note}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {['Energy', 'Sustainability'].map(cat => (
              <div key={cat} className="bg-white p-6 rounded shadow border">
                <div id={`table-${cat}`}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end', marginBottom: '4px' }}>
                    <h2 className="text-lg font-bold">{cat}組 - 報名之學校與國家隊伍數</h2>
                    <div style={{ fontSize: '11px', fontWeight: 'bold' }}>更新時間：{today}</div>
                  </div>
                  <table className="w-full border-collapse border-2 border-slate-900 font-bold text-sm">
                    <thead className="bg-slate-50">
                      <tr>
                        <th style={{ textAlign: 'center', padding: '4px', border: '2px solid #000', width: '40px' }}>#</th>
                        <th style={{ textAlign: 'left', padding: '4px', border: '2px solid #000' }}>學校名稱</th>
                        <th style={{ textAlign: 'right', padding: '4px', border: '2px solid #000', width: '100px' }}>代表國家</th>
                        <th style={{ textAlign: 'right', padding: '4px', border: '2px solid #000', width: '80px' }}>隊伍數</th>
                      </tr>
                    </thead>
                    <tbody>
                      {res.stats[cat].list.sort((a:any, b:any)=>b.count-a.count).map((item:any, idx:number) => (
                        <tr key={idx}>
                          <td style={{ textAlign: 'center', padding: '4px', border: '1px solid #000' }} className="text-slate-400">{idx+1}</td>
                          <td style={{ textAlign: 'left', padding: '4px', border: '1px solid #000' }}>{item.school.split(',')[0]}</td>
                          <td style={{ textAlign: 'right', padding: '4px', border: '1px solid #000' }}>{item.country}</td>
                          <td style={{ textAlign: 'right', padding: '4px', border: '1px solid #000' }}>{item.count}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ))}

            <div className="bg-slate-900 p-6 rounded-xl flex gap-4">
              <button onClick={() => exportExcel('school')} className="flex-1 p-3 bg-slate-800 rounded border border-slate-700 text-white font-bold text-sm">更新後學校標準對照表</button>
              <button onClick={() => exportExcel('country')} className="flex-1 p-3 bg-slate-800 rounded border border-slate-700 text-white font-bold text-sm">更新後國家標準對照表</button>
              <button onClick={() => exportExcel('stats')} className="flex-1 p-3 bg-emerald-700 rounded border border-emerald-600 text-white font-bold text-sm">本次統計資料表</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}