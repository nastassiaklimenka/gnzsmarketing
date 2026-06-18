import fs from "fs";

const SEO_PATH = "D:\\Генезис\\Marketing\\gnzsmarketing\\SEO";

export function saveToCSV(query, data) {
  const safeName = query.replace(/[^a-zа-я0-9]/gi, "_");

  const filePath = `${SEO_PATH}\\${safeName}.csv`;

  const csv = [
    "keyword;frequency",
    ...data.map((x) => `${x.keyword};${x.frequency}`),
  ].join("\n");

  fs.writeFileSync(filePath, csv, "utf-8");

  return filePath;
}
