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

  // 強化字串清理：轉小寫、去空白、統一國家名稱判定
  const cleanStr = (s: any) => String(s || "").replace(/\s+/g, '').toLowerCase();
  const today = new Date().toLocaleDateString();

  const res = useMemo(() => {
    if (data.length === 0) return null;
    const teams: { [key: string]: any } = {};
    const unSchools = new Set<string>();
    const currentFullSMap = { ...sMap, ...fixes };

    data.forEach(row => {
      const sn = String(row["隊伍序號"] || "").trim();
      if (sn.startsWith("888") || !sn.startsWith("99")) return;
      const [gid, mid] = sn.split("_");
      if (!teams[gid]) {
        teams[gid] = { 
          category: String(row["賽別"] || "").trim(), 
          members: [], 
          school: String(row["學校"] || "").trim() 
        };
      }
      teams[gid].members.push(parseInt(mid));
    });

    const stats: any = {
      Energy: { teams: 0, people: 0, schools: new Set(), countries: new Set(), overseas: 0, list: [] },
      Sustainability: { teams: 0, people: 0, schools: new Set(), countries: new Set(), overseas: 0, list: [] }
    };

    Object.keys(teams).forEach(gid => {
      const t = teams[gid];
      const cat = t.category.toLowerCase().includes("energy") ? "Energy" : "Sustainability";
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
        const rawC = cleanStr(parts[parts.length - 1]);
        // 解決 Phillippines 判定問題
        const countryMatchKey = Object.keys(cMap).find(k => cleanStr(k).includes(rawC) || rawC.includes(cleanStr(k)));
        country = countryMatchKey ? cMap[countryMatchKey] : parts[parts.length - 1].trim();
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

  const handleDownloadPNG = (id: string, name: string) => {
    const el = document.getElementById(id);
    if (el) {
      html2canvas(el, { scale: 3 }).then(canvas => {
        const link = document.createElement('a');
        link.download = `${name}_${today}.png`;
        link.href = canvas.toDataURL();
        link.click();
      });
    }
  };

  const exportExcelTotal = () => {
    if (!res) return;
    const wb = XLSX.utils.book_new();

    // 1. Summary Sheet
    const summaryData = [
      ["統計項目", "Energy", "Sustainability", "合計", "備註"],
      ["報名隊伍數", res.stats.Energy.teams, res.stats.Sustainability.teams, res.stats.Energy.teams + res.stats.Sustainability.teams, ""],
      ["報名人數", res.stats.Energy.people, res.stats.Sustainability.people, res.stats.Energy.people + res.stats.Sustainability.people, ""],
      ["報名學校數", res.stats.Energy.schools.size, res.stats.Sustainability.schools.size, new Set([...res.stats.Energy.schools, ...res.stats.Sustainability.schools]).size, "不計算重複學校"],
      ["海外學校數", res.stats.Energy.overseas, res.stats.Sustainability.overseas, res.stats.Energy.overseas + res.stats.Sustainability.overseas, ""],
      ["報名國家數", res.stats.Energy.countries.size, res.stats.Sustainability.countries.size, new Set([...res.stats.Energy.countries, ...res.stats.Sustainability.countries]).size, "不計算重複國家"]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), "統計總表");

    // 2. Energy Sheet
    const energyData = [["#", "學校名稱", "代表國家", "隊伍數"], ...res.stats.Energy.list.sort((a:any,b:any)=>b.count-a.count).map((item:any, i:number)=>[i+1, item.school, item.country, item.count])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(energyData), "Energy名單");

    // 3. Sust Sheet
    const sustData = [["#", "學校名稱", "代表國家", "隊伍數"], ...res.stats.Sustainability.list.sort((a:any,b:any)=>b.count-a.count).map((item:any, i:number)=>[i+1, item.school, item.country, item.count])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sustData), "Sustainability名單");

    XLSX.writeFile(wb, `2026_NZ_完整報表_${today}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-slate-50 pb-40" style={{ fontFamily: '"Microsoft JhengHei", sans-serif' }}>
      {/* Banner 區域 */}
      <div className="w-full bg-slate-200 flex items-center justify-center relative shadow-sm" style={{ minHeight: bannerLoaded ? 'auto' : '150px' }}>
        <img src="/banner.png" alt="Banner" className="w-full h-auto block" onLoad={() => setBannerLoaded(true)} onError={(e) => { e.currentTarget.style.display = 'none'; setBannerLoaded(false); }} />
        {!bannerLoaded && <p className="absolute text-slate-400 font-bold text-xl">Banner 區域 (請放入 banner.png)</p>}
      </div>

      <div className="max-w-[1200px] mx-auto mt-12 px-6">
        <div className="bg-white p-10 rounded-3xl shadow-xl border border-slate-200 mb-16">
          <div className="flex justify-between items-center mb-10">
            <h1 className="text-4xl font-black border-l-8 border-slate-900 pl-6 uppercase">2026 NZ 統計控制中心</h1>
            {res && (
              <button onClick={exportExcelTotal} className="bg-emerald-600 text-white px-8 py-3 rounded-xl font-black hover:bg-emerald-700 transition shadow-lg">匯出完整 EXCEL 報表</button>
            )}
          </div>
          <div className="grid grid-cols-3 gap-8">
            {[{ t: '1. 原始報名資料', id: 'raw' }, { t: '2. 學校標準表', id: 'school' }, { t: '3. 國家標準表', id: 'country' }].map(box => (
              <div key={box.id} className="p-6 border-2 border-slate-100 rounded-2xl bg-slate-50/50">
                <p className="font-black text-lg mb-3">{box.t}</p>
                <input type="file" onChange={e => loadFile(e, box.id)} className="text-xs w-full"/>
              </div>
            ))}
          </div>
        </div>

        {res && (
          <div className="space-y-24">
            {/* Report 01: 總表 */}
            <div className="bg-white p-12 rounded-3xl shadow-2xl border border-slate-200">
              <div className="flex justify-between items-center mb-8">
                <h4 className="text-2xl font-black bg-slate-900 text-white px-6 py-2 rounded-full">Report 01</h4>
                <button onClick={() => handleDownloadPNG('table-summary', 'NZ目前報名狀況統計表')} className="bg-slate-800 text-white px-6 py-2 rounded-lg font-bold">下載 PNG</button>
              </div>
              <div id="table-summary" className="p-8 bg-white border-[6px] border-slate-900">
                <div className="flex justify-between items-end mb-6 border-b-4 border-slate-900 pb-2 font-black">
                  <h2 className="text-4xl">NZ 目前報名狀況統計表</h2>
                  <div className="text-right">資料更新日期：{today}</div>
                </div>
                <table className="w-full border-collapse border-[4px] border-slate-900 text-2xl font-black">
                  <thead className="bg-slate-100">
                    <tr className="border-b-[4px] border-slate-900">
                      <th className="border-r-4 border-slate-900 p-4 text-left">統計項目</th>
                      <th className="border-r-4 border-slate-900 p-4 text-right">Energy</th>
                      <th className="border-r-4 border-slate-900 p-4 text-right">Sustainability</th>
                      <th className="border-r-4 border-slate-900 p-4 text-right bg-slate-200">合計</th>
                      <th className="p-4 text-left text-base text-slate-500">備註</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[
                      { l: '報名隊伍數', k: 'teams' }, { l: '報名人數', k: 'people' }, 
                      { l: '報名學校數', k: 'schools', note: '不計算重複學校' }, 
                      { l: '海外學校數', k: 'overseas', c: 'text-rose-600' }, 
                      { l: '報名國家數', k: 'countries', note: '不計算重複國家' }
                    ].map((row, i) => (
                      <tr key={i} className="border-b-4 border-slate-900">
                        <td className="border-r-4 border-slate-900 p-4 text-left bg-slate-50">{row.l}</td>
                        <td className={`border-r-4 border-slate-900 p-4 text-right ${row.c}`}>{i === 2 || i === 4 ? res.stats.Energy[row.k].size : res.stats.Energy[row.k]}</td>
                        <td className={`border-r-4 border-slate-900 p-4 text-right ${row.c}`}>{i === 2 || i === 4 ? res.stats.Sustainability[row.k].size : res.stats.Sustainability[row.k]}</td>
                        <td className={`border-r-4 border-slate-900 p-4 text-right bg-slate-100 ${row.c}`}>
                          {row.k === 'teams' ? (res.stats.Energy.teams + res.stats.Sustainability.teams) :
                           row.k === 'people' ? (res.stats.Energy.people + res.stats.Sustainability.people) :
                           row.k === 'overseas' ? (res.stats.Energy.overseas + res.stats.Sustainability.overseas) :
                           row.k === 'schools' ? new Set([...res.stats.Energy.schools, ...res.stats.Sustainability.schools]).size :
                           new Set([...res.stats.Energy.countries, ...res.stats.Sustainability.countries]).size}
                        </td>
                        <td className="p-4 text-left text-sm text-slate-400 font-normal">{row.note}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Energy & Sustainability 分表... (省略部分重複樣式，已確保靠右對齊) */}
            <div className="grid grid-cols-1 gap-24">
               {['Energy', 'Sustainability'].map(cat => (
                 <div key={cat} className="bg-white p-12 rounded-3xl shadow-2xl border border-slate-200">
                    <div className="flex justify-between items-center mb-8">
                      <h4 className={`text-2xl font-black ${cat==='Energy'?'bg-blue-600':'bg-emerald-600'} text-white px-6 py-2 rounded-full`}>{cat} 組報名清單</h4>
                      <button onClick={() => handleDownloadPNG(`table-${cat}`, `${cat}組-報名名單`)} className="bg-slate-800 text-white px-6 py-2 rounded-lg font-bold">下載 PNG</button>
                    </div>
                    <div id={`table-${cat}`} className={`p-8 bg-white border-[6px] ${cat==='Energy'?'border-blue-600':'border-emerald-600'}`}>
                      <h2 className={`text-4xl font-black mb-6 border-b-4 ${cat==='Energy'?'border-blue-600 text-blue-600':'border-emerald-600 text-emerald-600'} pb-2`}>{cat}組 - 報名之學校與國家隊伍數</h2>
                      <table className="w-full text-xl border-collapse border-[4px] border-slate-900 font-black">
                        <thead className="bg-slate-100">
                          <tr className="border-b-4 border-slate-900">
                            <th className="border-r-4 border-slate-900 p-3 w-16 text-center">#</th>
                            <th className="border-r-4 border-slate-900 p-3 text-left">學校名稱</th>
                            <th className="border-r-4 border-slate-900 p-3 w-40 text-right">代表國家</th>
                            <th className="p-3 w-32 text-right">隊伍數</th>
                          </tr>
                        </thead>
                        <tbody>
                          {res.stats[cat].list.sort((a:any, b:any)=>b.count - a.count).map((item:any, idx:number) => (
                            <tr key={idx} className="border-b-2 border-slate-200">
                              <td className="border-r-2 p-3 text-center text-slate-300">{idx+1}</td>
                              <td className="border-r-2 p-3 text-left">{item.school.split(',')[0]}</td>
                              <td className="border-r-2 p-3 text-right">{item.country}</td>
                              <td className={`p-3 text-right text-2xl ${cat==='Energy'?'text-blue-600':'text-emerald-600'}`}>{item.count}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                 </div>
               ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}