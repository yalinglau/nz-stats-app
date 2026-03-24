import { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';

type StandardMap = { [key: string]: string };

export default function App() {
  const [data, setData] = useState<any[]>([]);
  const [sMap, setSMap] = useState<StandardMap>({});
  const [cMap, setCMap] = useState<StandardMap>({});
  const [fixes, setFixes] = useState<StandardMap>({});
  const [tempFixes, setTempFixes] = useState<StandardMap>({});

  const todayStr = `${new Date().getMonth() + 1}/${new Date().getDate()}日`;
  const reportTitle = `NZ ${todayStr} 報名結果`;

  const cleanStr = (s: any) => String(s || "").replace(/\s+/g, '').toLowerCase();

  const res = useMemo(() => {
    if (data.length === 0) return null;
    const teams: { [key: string]: any } = {};
    const unSchools = new Set<string>();
    const currentFullSMap = { ...sMap, ...fixes };

    data.forEach(row => {
      const sn = String(row["隊伍序號"] || "").trim();
      if (!sn || sn.startsWith("888")) return;
      const [gid, mid] = sn.split("_");
      if (!teams[gid]) {
        teams[gid] = { 
          category: String(row["賽別"] || "").trim(),
          school: String(row["學校"] || "").trim(),
          members: []
        };
      }
      teams[gid].members.push(mid ? parseInt(mid) : 0);
    });

    const stats: any = {
      Energy: { teams: 0, people: 0, schools: new Set(), twSchools: new Set(), countries: new Set(), overseas: 0, list: [] },
      Sustainability: { teams: 0, people: 0, schools: new Set(), twSchools: new Set(), countries: new Set(), overseas: 0, list: [] }
    };

    Object.keys(teams).forEach(gid => {
      const t = teams[gid];
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
      let isOverseas = false;

      if (stdSchool.includes(",")) {
        isOverseas = true;
        const parts = stdSchool.split(",");
        const rawC = parts[parts.length - 1].trim();
        const countryMatchKey = Object.keys(cMap).find(k => cleanStr(k) === cleanStr(rawC));
        country = countryMatchKey ? cMap[countryMatchKey] : rawC;
      }

      stats[cat].teams += 1;
      stats[cat].people += pCount;
      stats[cat].schools.add(stdSchool);
      stats[cat].countries.add(country);
      
      if (isOverseas) stats[cat].overseas += 1;
      else stats[cat].twSchools.add(stdSchool);
      
      const existing = stats[cat].list.find((i:any) => i.school === stdSchool);
      if (existing) existing.count += 1;
      else stats[cat].list.push({ school: stdSchool, country, count: 1 });
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
      const m: StandardMap = {};
      json.forEach((row: any) => {
        const ks = Object.keys(row);
        if (row[ks[0]]) m[String(row[ks[0]]).trim()] = String(row[ks[1]] || "").trim();
      });
      if (type === 'raw') setData(json);
      else if (type === 'school') setSMap(m);
      else setCMap(m);
    };
    r.readAsBinaryString(f);
  };

  const applyFix = (raw: string) => {
    if (tempFixes[raw]) {
      setFixes(prev => ({ ...prev, [raw]: tempFixes[raw] }));
    }
  };

  const exportExcel = () => {
    if (!res) return;
    const wb = XLSX.utils.book_new();
    const rowsConfig = [
        { label: '報名隊伍數', key: 'teams', note: '' },
        { label: '報名人數', key: 'people', note: '' },
        { label: '報名學校數', key: 'schools', note: '不計算重複學校' },
        { label: '- 臺灣學校數', key: 'twSchools', note: '不計算重複學校' },
        { label: '- 海外學校數', key: 'overseas', note: '' },
        { label: '報名國家數', key: 'countries', note: '含臺灣，不計算重複國家' }
    ];

    const s1Head = [["項目", "Energy", "Sustainability", "合計", "備註"]];
    const s1Body = rowsConfig.map(r => {
        const vE = (res.stats.Energy[r.key] instanceof Set) ? res.stats.Energy[r.key].size : res.stats.Energy[r.key];
        const vS = (res.stats.Sustainability[r.key] instanceof Set) ? res.stats.Sustainability[r.key].size : res.stats.Sustainability[r.key];
        const total = (res.stats.Energy[r.key] instanceof Set) 
            ? new Set([...res.stats.Energy[r.key], ...res.stats.Sustainability[r.key]]).size 
            : vE + vS;
        return [r.label, vE, vS, total, r.note];
    });
    
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([...s1Head, ...s1Body]), "統計總表");

    ['Energy', 'Sustainability'].forEach(cat => {
        const sHead = [[`報名學校與國家 - ${cat}`], ["序", "學校名稱", "代表國家", "隊伍數"]];
        const sBody = res.stats[cat].list.sort((a:any, b:any) => b.count - a.count).map((item:any, i:number) => [i + 1, item.school, item.country, item.count]);
        const sFooter = [["合計", "", "", res.stats[cat].teams]];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([...sHead, ...sBody, ...sFooter]), `${cat}組`);
    });
    XLSX.writeFile(wb, `${reportTitle}.xlsx`);
  };

  const exportMap = (type: 'school' | 'country') => {
    const wb = XLSX.utils.book_new();
    const current = type === 'school' ? { ...sMap, ...fixes } : cMap;
    const exportData = Object.entries(current).map(([k, v]) => [k, v]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(exportData), "Sheet1");
    XLSX.writeFile(wb, `更新後${type === 'school' ? '學校' : '國家'}標準對照表.xlsx`);
  };

  return (
    <div className="min-h-screen bg-white font-sans text-slate-900">
      {/* 1. Banner (讀取 public/banner.png) */}
      <div className="w-full">
        <img src="/banner.png" alt="NZ Banner" className="w-full h-auto block" />
      </div>

      <div className="max-w-[1100px] mx-auto px-6 pb-20 mt-8">
        <h2 className="text-2xl font-bold mb-6">2026 NZ 統計控制中心</h2>
        
        <div className="space-y-6 mb-12 border-b pb-8">
          <div><p className="font-bold">1. 原始報名資料</p><input type="file" onChange={e => loadFile(e, 'raw')} className="mt-2 text-sm" /></div>
          <div><p className="font-bold">2. 學校標準表</p><input type="file" onChange={e => loadFile(e, 'school')} className="mt-2 text-sm" /></div>
          <div><p className="font-bold">3. 國家標準表</p><input type="file" onChange={e => loadFile(e, 'country')} className="mt-2 text-sm" /></div>
        </div>

        {res && res.unSchools.length > 0 && (
          <div className="mb-12 p-6 border-2 border-amber-200 bg-amber-50 rounded-lg">
            <p className="font-bold text-amber-900 mb-4 text-lg underline">未定義學校修正：</p>
            {res.unSchools.map(u => (
              <div key={u} className="flex gap-4 mb-2">
                <span className="bg-white p-2 border w-72 truncate font-mono text-sm">{u}</span>
                <input className="border p-2 flex-1 rounded text-sm" placeholder="輸入標準名稱" onChange={e => setTempFixes(prev => ({ ...prev, [u]: e.target.value }))} />
                <button onClick={() => applyFix(u)} className="bg-amber-600 text-white px-6 py-2 rounded font-bold">更新</button>
              </div>
            ))}
          </div>
        )}

        {res && (
          <div className="space-y-16">
            <div>
              <h2 className="text-2xl font-bold mb-6">{reportTitle}</h2>
              <table className="w-full border-collapse border-2 border-slate-300">
                <thead className="bg-slate-50 font-bold">
                  <tr>
                    <th className="border border-slate-300 p-3 text-left">項目</th>
                    <th className="border border-slate-300 p-3 text-right">Energy</th>
                    <th className="border border-slate-300 p-3 text-right">Sustainability</th>
                    <th className="border border-slate-300 p-3 text-right bg-slate-100">合計</th>
                    <th className="border border-slate-300 p-3 text-left">備註</th>
                  </tr>
                </thead>
                <tbody className="font-bold">
                  {[
                    { label: '報名隊伍數', key: 'teams', note: '' },
                    { label: '報名人數', key: 'people', note: '' },
                    { label: '報名學校數', key: 'schools', note: '不計算重複學校' },
                    { label: '- 臺灣學校數', key: 'twSchools', note: '不計算重複學校', indent: true },
                    { label: '- 海外學校數', key: 'overseas', note: '', indent: true },
                    { label: '報名國家數', key: 'countries', note: '含臺灣，不計算重複國家' }
                  ].map((row, i) => {
                    const vE = (res.stats.Energy[row.key] instanceof Set) ? res.stats.Energy[row.key].size : res.stats.Energy[row.key];
                    const vS = (res.stats.Sustainability[row.key] instanceof Set) ? res.stats.Sustainability[row.key].size : res.stats.Sustainability[row.key];
                    const total = (res.stats.Energy[row.key] instanceof Set) ? new Set([...res.stats.Energy[row.key], ...res.stats.Sustainability[row.key]]).size : vE + vS;
                    return (
                      <tr key={i} className="hover:bg-slate-50">
                        <td className={`border border-slate-300 p-3 ${row.indent ? 'pl-8 text-slate-600 font-medium' : ''}`}>{row.label}</td>
                        <td className="border border-slate-300 p-3 text-right">{vE}</td>
                        <td className="border border-slate-300 p-3 text-right">{vS}</td>
                        <td className="border border-slate-300 p-3 text-right bg-slate-50">{total}</td>
                        <td className="border border-slate-300 p-3 text-xs text-slate-500 font-normal">{row.note}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {['Energy', 'Sustainability'].map(cat => (
              <div key={cat}>
                <h2 className="text-xl font-bold mb-6">報名學校與國家 - {cat}</h2>
                <table className="w-full border-collapse border-2 border-slate-300">
                  <thead className="bg-slate-50 text-sm">
                    <tr><th className="border border-slate-300 p-3 w-20">序</th><th className="border border-slate-300 p-3 text-left">學校名稱</th><th className="border border-slate-300 p-3 text-left">代表國家</th><th className="border border-slate-300 p-3 text-right w-32">隊伍數</th></tr>
                  </thead>
                  <tbody>
                    {res.stats[cat].list.sort((a:any, b:any) => b.count - a.count).map((item:any, i:number) => (
                      <tr key={i}><td className="border border-slate-300 p-3 text-center text-slate-400">{i + 1}</td><td className="border border-slate-300 p-3 font-bold">{item.school}</td><td className="border border-slate-300 p-3">{item.country}</td><td className="border border-slate-300 p-3 text-right font-bold">{item.count}</td></tr>
                    ))}
                    <tr className="bg-slate-100 font-black"><td className="border border-slate-300 p-3 text-center">合計</td><td className="border border-slate-300 p-3"></td><td className="border border-slate-300 p-3"></td><td className="border border-slate-300 p-3 text-right">{res.stats[cat].teams}</td></tr>
                  </tbody>
                </table>
              </div>
            ))}

            <div className="flex justify-center gap-6 pt-12 border-t-2 border-slate-100">
              <button onClick={() => exportMap('school')} className="px-8 py-4 border-2 border-slate-800 rounded bg-slate-900 text-white font-bold hover:bg-slate-800 transition">更新後學校標準對照表</button>
              <button onClick={() => exportMap('country')} className="px-8 py-4 border-2 border-slate-800 rounded bg-slate-900 text-white font-bold hover:bg-slate-800 transition">更新後國家標準對照表</button>
              <button onClick={exportExcel} className="px-12 py-4 rounded bg-emerald-700 text-white font-black text-lg shadow-lg hover:bg-emerald-800 transition">本次統計資料表</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}