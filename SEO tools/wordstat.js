import { chromium } from "playwright";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const COOKIES_PATH = path.join(__dirname, "cookies.json");

export async function getWordstat(query) {
  const browser = await chromium.launch({
    headless: false,
  });

  const context = await browser.newContext();

  // 🍪 подгружаем cookies
  if (fs.existsSync(COOKIES_PATH)) {
    const cookies = JSON.parse(fs.readFileSync(COOKIES_PATH, "utf-8"));
    await context.addCookies(cookies);
  }

  const page = await context.newPage();

  await page.goto("https://wordstat.yandex.ru");

  // ждём загрузку
  await page.waitForTimeout(3000);

  // 🔍 ввод запроса
  await page.keyboard.type(query);
  await page.keyboard.press("Enter");

  // ⏳ ждём таблицу результатов
  await page.waitForTimeout(6000);

  // 📊 вытаскиваем данные из таблицы
  const data = await page.evaluate(() => {
    const rows = Array.from(document.querySelectorAll("table tbody tr"));

    return rows
      .map((row) => {
        const cols = row.querySelectorAll("td");

        return {
          keyword: cols[0]?.innerText?.trim(),
          frequency: cols[1]?.innerText?.trim(),
        };
      })
      .filter((x) => x.keyword);
  });

  await browser.close();

  return {
    query,
    count: data.length,
    data,
  };
}
