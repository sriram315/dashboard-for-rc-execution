
import { DashboardData } from "../types";

export const parseCSV = (csvText: string): DashboardData => {
  const lines = csvText.split(/\r?\n/).filter(line => line.trim() !== "");
  if (lines.length === 0) return { headers: [], rows: [] };

  // Improved regex to handle commas inside quotes
  const splitLine = (line: string) => {
    const result = [];
    let cur = "";
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === ',' && !inQuotes) {
        result.push(cur.trim());
        cur = "";
      } else {
        cur += char;
      }
    }
    result.push(cur.trim());
    return result;
  };

  const headers = splitLine(lines[0]);
  const rows = lines.slice(1).map(line => {
    const values = splitLine(line);
    const row: Record<string, any> = {};
    
    headers.forEach((header, index) => {
      let val: any = values[index] || "";
      
      // Clean up surrounding quotes if they exist
      if (typeof val === 'string') {
        val = val.replace(/^"|"$/g, '');
      }

      // Check if value is numeric
      if (val !== "" && !isNaN(val as any) && typeof val !== 'boolean') {
        row[header] = Number(val);
      } else {
        row[header] = val;
      }
    });
    return row;
  });

  return { headers, rows };
};
