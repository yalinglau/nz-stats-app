import React, { useState, useMemo } from 'react';
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

      const pCount = (t.members.length === 1 && t.members[0] === 0) ? 1 : t.members.filter(m => m !== 0).length;
      let country = "臺灣";
      if (stdSchool.includes(",")) {
        const parts = stdSchool.split(",");
        const rawC = cleanStr(parts[parts.length - 1]);
        const countryMatchKey = Object.keys(cMap).find(k => cleanStr(k) === rawC);
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

  const handleDownload = (id: string, name: string) => {
    const el = document.getElementById(id);
    if (el) {
      html2canvas(el, { scale: 3 }).then(canvas => {
        const link = document.createElement('a');
        link.download = `${name}_${new Date().toISOString().slice(0,10)}.png`;
        link.href = canvas.toDataURL();
        link.click();
      });
    }
  };

  const exportMap = (type: 'school' | 'country') => {
    const targetMap = type === 'school' ? { ...sMap, ...fixes } : cMap;
    const fileName = type === 'school' ? "最新學校對照表" : "最新國家對照表";
    const ws = XLSX.utils.json_to_sheet(Object.entries(targetMap).map(([k, v]) => ({ "原始名稱": k, "標準名稱": v })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, `2026_NZ_${fileName}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-slate-50 pb-40" style={{ fontFamily: '"Microsoft JhengHei", "微軟正黑體", Arial, sans-serif' }}>
      
      {/* Banner 區 */}
      <div className="w-full bg-slate-200 flex items-center justify-center relative shadow-sm" style={{ minHeight: bannerLoaded ? 'auto' : '150px' }}>
        <img src="/banner.png" alt="Banner" className="w-full h-auto block" onLoad={() => setBannerLoaded(true)} onError={(e) => { e.currentTarget.style.display = 'none'; setBannerLoaded(false); }} />
        {!bannerLoaded && <p className="absolute text-slate-400 font-bold text-xl uppercase tracking-widest">Banner 區域</p>}
      </div>

      <div className="max-w-[1200px] mx-auto mt-12 px-6">
        
        {/* 1. 控制台與上傳區 */}
        <div className="bg-white p-10 rounded-3xl shadow-xl border border-slate-200 mb-16">
          <h1 className="text-4xl font-black mb-10 border-l-8 border-slate-900 pl-6 uppercase">2026 NZ 統計控制中心</h1>
          <div className="grid grid-cols-3 gap-8">
            {[{ t: '原始報名資料', id: 'raw', c: 'blue' }, { t: '學校標準表', id: 'school', c: 'slate' }, { t: '國家標準表', id: 'country', c: 'slate' }].map(box => (
              <div key={box.id} className="p-6 border-2 border-slate-100 rounded-2xl hover:border-blue-200 transition bg-slate-50/50">
                <p className="font-black text-lg mb-3">{box.t}</p>
                <input type="file" onChange={e => loadFile(e, box.id)} className="text-xs w-full cursor-pointer"/>
              </div>
            ))}
          </div>
        </div>

        {/* 2. 校名修正區 */}
        {res && res.unSchools.length > 0 && (
          <div className="mb-16 p-10 bg-rose-50 border-4 border-rose-200 rounded-3xl shadow-xl">
            <div className="flex justify-between items-center mb-8">
              <h3 className="font-black text-rose-800 text-3xl">⚠️ 偵測到 {res.unSchools.length} 個未知校名</h3>
              <button onClick={() => {setFixes(prev => ({...prev, ...tempFixes})); setTempFixes({});}} className="bg-rose-600 text-white px-10 py-4 rounded-2xl font-black text-xl hover:scale-105 transition shadow-xl">更新並重新統計</button>
            </div>
            <div className="grid grid-cols-1 gap-4 max-h-64 overflow-y-auto pr-4 font-bold">
              {res.unSchools.map((us) => (
                <div key={us} className="flex items-center gap-6 bg-white p-5 border-2 rounded-2xl shadow-sm">
                  <span className="text-sm text-slate-400 w-1/3 truncate">{us}</span>
                  <input className="flex-1 border-b-4 border-slate-100 px-4 py-2 text-xl outline-none focus:border-rose-500" placeholder="填寫標準校名" onChange={e => setTempFixes({...tempFixes, [us]: e.target.value})}/>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* 3. 統計結果展現區 (分為三個卡片，確保間距) */}
        {res && (
          <div className="space-y-24"> {/* 強大的垂直間距 */}
            
            {/* (1) NZ目前報名狀況統計表 */}
            <div className="bg-white p-12 rounded-3xl shadow-2xl border border-slate-200">
              <div className="flex justify-between items-center mb-8">
                <h4 className="text-2xl font-black bg-slate-900 text-white px-6 py-2 rounded-full">Report 01</h4>
                <button onClick={() => handleDownload('table-summary', 'NZ目前報名狀況統計表')} className="bg-slate-800 text-white px-6 py-2 rounded-lg font-bold hover:invert transition">下載 PNG</button>
              </div>
              <div id="table-summary" className="p-8 bg-white border-[6px] border-slate-900">
                <div className="flex justify-between items-end mb-6 border-b-4 border-slate-900 pb-2">
                  <h2 className="text-4xl font-black">NZ 目前報名狀況統計表</h2>
                  <div className="text-right font-black text-lg">資料更新日期：{today}</div>
                </div>
                <table className="w-full border-collapse border-[4px] border-slate-900 text-2xl text-center font-black">
                  <thead className="bg-slate-100">
                    <tr className="border-b-[4px] border-slate-900">
                      <th className="border-r-4 border-slate-900 p-4 text-left">統計項目</th>
                      <th className="border-r-4 border-slate-900 p-4">Energy</th>
                      <th className="border-r-4 border-slate-900 p-4">Sustainability</th>
                      <th className="p-4 bg-slate-200">合計</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[
                      { l: '報名隊伍數', k: 'teams' }, { l: '報名人數', k: 'people' }, 
                      { l: '報名學校數', k: 'schools' }, { l: '海外學校數', k: 'overseas', c: 'text-rose-600' }, 
                      { l: '報名國家數', k: 'countries' }
                    ].map((row, i) => (
                      <tr key={i} className="border-b-4 border-slate-900">
                        <td className="border-r-4 border-slate-900 p-4 text-left bg-slate-50 whitespace-nowrap">{row.l}</td>
                        <td className={`border-r-4 border-slate-900 p-4 ${row.c}`}>{i === 2 || i === 4 ? res.stats.Energy[row.k].size : res.stats.Energy[row.k]}</td>
                        <td className={`border-r-4 border-slate-900 p-4 ${row.c}`}>{i === 2 || i === 4 ? res.stats.Sustainability[row.k].size : res.stats.Sustainability[row.k]}</td>
                        <td className={`p-4 bg-slate-100 ${row.c}`}>
                          {row.k === 'teams' ? (res.stats.Energy.teams + res.stats.Sustainability.teams) :
                           row.k === 'people' ? (res.stats.Energy.people + res.stats.Sustainability.people) :
                           row.k === 'overseas' ? (res.stats.Energy.overseas + res.stats.Sustainability.overseas) :
                           row.k === 'schools' ? new Set([...res.stats.Energy.schools, ...res.stats.Sustainability.schools]).size :
                           new Set([...res.stats.Energy.countries, ...res.stats.Sustainability.countries]).size}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* (2) Energy 分表 */}
            <div className="bg-white p-12 rounded-3xl shadow-2xl border border-slate-200">
              <div className="flex justify-between items-center mb-8">
                <h4 className="text-2xl font-black bg-blue-600 text-white px-6 py-2 rounded-full">Report 02</h4>
                <button onClick={() => handleDownload('table-energy', 'Energy組-報名名單')} className="bg-blue-600 text-white px-6 py-2 rounded-lg font-bold hover:invert transition">下載 PNG</button>
              </div>
              <div id="table-energy" className="p-8 bg-white border-[6px] border-blue-600">
                <div className="flex justify-between items-end mb-6 border-b-4 border-blue-600 pb-2 font-black text-blue-600">
                  <h2 className="text-4xl uppercase">Energy組 - 報名之學校與國家隊伍數</h2>
                  <div className="text-lg font-black">更新時間：{today}</div>
                </div>
                <table className="w-full text-xl border-collapse border-[4px] border-slate-900 font-black">
                  <thead className="bg-slate-100 text-center">
                    <tr className="border-b-4 border-slate-900">
                      <th className="border-r-4 border-slate-900 p-3 w-16">#</th>
                      <th className="border-r-4 border-slate-900 p-3 text-left">學校名稱</th>
                      <th className="border-r-4 border-slate-900 p-3 w-32">代表國家</th>
                      <th className="p-3 w-24">隊伍數</th>
                    </tr>
                  </thead>
                  <tbody>
                    {res.stats.Energy.list.sort((a:any, b:any)=>b.count - a.count).map((item:any, idx:number) => (
                      <tr key={idx} className="border-b-2 border-slate-200 text-center">
                        <td className="border-r-2 p-3 text-slate-300 font-mono text-base">{idx+1}</td>
                        <td className="border-r-2 p-3 text-left leading-tight">{item.school.split(',')[0]}</td>
                        <td className="border-r-2 p-3">{item.country}</td>
                        <td className="p-3 text-blue-600 text-2xl font-black">{item.count}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* (3) Sustainability 分表 */}
            <div className="bg-white p-12 rounded-3xl shadow-2xl border border-slate-200">
              <div className="flex justify-between items-center mb-8">
                <h4 className="text-2xl font-black bg-emerald-600 text-white px-6 py-2 rounded-full">Report 03</h4>
                <button onClick={() => handleDownload('table-sust', 'Sustainability組-報名名單')} className="bg-emerald-600 text-white px-6 py-2 rounded-lg font-bold hover:invert transition">下載 PNG</button>
              </div>
              <div id="table-sust" className="p-8 bg-white border-[6px] border-emerald-600">
                <div className="flex justify-between items-end mb-6 border-b-4 border-emerald-600 pb-2 font-black text-emerald-600">
                  <h2 className="text-4xl uppercase">Sustainability組 - 報名之學校與國家隊伍數</h2>
                  <div className="text-lg font-black">更新時間：{today}</div>
                </div>
                <table className="w-full text-xl border-collapse border-[4px] border-slate-900 font-black">
                  <thead className="bg-slate-100 text-center">
                    <tr className="border-b-4 border-slate-900">
                      <th className="border-r-4 border-slate-900 p-3 w-16">#</th>
                      <th className="border-r-4 border-slate-900 p-3 text-left">學校名稱</th>
                      <th className="border-r-4 border-slate-900 p-3 w-32">代表國家</th>
                      <th className="p-3 w-24">隊伍數</th>
                    </tr>
                  </thead>
                  <tbody>
                    {res.stats.Sustainability.list.sort((a:any, b:any)=>b.count - a.count).map((item:any, idx:number) => (
                      <tr key={idx} className="border-b-2 border-slate-200 text-center">
                        <td className="border-r-2 p-3 text-slate-300 font-mono text-base">{idx+1}</td>
                        <td className="border-r-2 p-3 text-left leading-tight">{item.school.split(',')[0]}</td>
                        <td className="border-r-2 p-3">{item.country}</td>
                        <td className="p-3 text-emerald-600 text-2xl font-black">{item.count}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* 4. 資料庫維護區 (移動至最下方) */}
            <div className="bg-slate-900 p-12 rounded-3xl shadow-2xl text-white">
              <div className="flex justify-between items-center mb-8 border-b border-slate-700 pb-6">
                <div>
                  <h2 className="text-3xl font-black mb-2 tracking-widest">資料維護中心</h2>
                  <p className="text-slate-400 font-bold">當前所有手動修正的資料均已同步至記憶體中</p>
                </div>
                <div className="flex gap-6">
                  <button onClick={() => exportMap('school')} className="bg-emerald-500 text-slate-900 px-8 py-3 rounded-xl font-black hover:bg-white transition">匯出最新學校對照表</button>
                  <button onClick={() => exportMap('country')} className="bg-blue-400 text-slate-900 px-8 py-3 rounded-xl font-black hover:bg-white transition">匯出最新國家對照表</button>
                </div>
              </div>
              <p className="text-center text-slate-600 font-bold italic">2026 NZ AUTO-STATS ENGINE v1.5 - Optimized for Individual Reporting</p>
            </div>

          </div>
        )}
      </div>
    </div>
  );
}