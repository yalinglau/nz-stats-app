import * as XLSX from 'xlsx';

// 使用 type 確保在 Vite 編譯時能被正確導出
export type StandardMap = { 
  [key: string]: string 
};

export const processNZData = (rawData: any[], schoolMap: StandardMap, countryMap: StandardMap) => {
  const teams: { [key: string]: { category: string, members: number[], school: string } } = {};
  const unSchools = new Set<string>();
  const unCountries = new Set<string>();

  // 1. 基礎分組與過濾 (排除 888 測試帳號，僅納入 99 開頭的序號)
  rawData.forEach(row => {
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

  // 2. 統計運算邏輯
  Object.keys(teams).forEach(gid => {
    const t = teams[gid];
    // 賽別判定 (相容 Energy 與 Sustainability 關鍵字)
    const cat = t.category.toLowerCase().includes("energy") ? "Energy" : "Sustainability";
    if (!stats[cat]) return;

    // 規則 2: 人數判定 (單人 _0 算 1 人，多人則過濾 _0)
    const pCount = (t.members.length === 1 && t.members[0] === 0) ? 1 : t.members.filter(m => m !== 0).length;

    // 規則 3: 學校對照
    const stdSchool = schoolMap[t.school];
    if (!stdSchool) {
      unSchools.add(t.school);
      return;
    }

    // 規則 4: 海外判定 (逗號標記法) 與 國家翻譯
    let country = "臺灣";
    if (stdSchool.includes(",")) {
      const parts = stdSchool.split(",");
      const rawC = parts[parts.length - 1].trim(); // 取最後一個逗號後面的內容
      country = countryMap[rawC] || rawC;
      if (!countryMap[rawC]) unCountries.add(rawC);
    }

    stats[cat].teams += 1;
    stats[cat].people += pCount;
    stats[cat].schools.add(stdSchool);
    stats[cat].countries.add(country);
    if (stdSchool.includes(",")) stats[cat].overseas += 1;
    stats[cat].list.push({ school: stdSchool, country, gid });
  });

  return { 
    stats, 
    unSchools: Array.from(unSchools), 
    unCountries: Array.from(unCountries) 
  };
};