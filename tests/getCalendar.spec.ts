import test, { Locator, Page } from "@playwright/test";
import { fullLoadingDone, getLocator, globalVariables } from "../helper/helper";
import {
  ClaimDay,
  ExcelData,
  generateMealExcel,
} from "../helper/excelGenerator";

const getHtmlStringInnerHTML = async (
  page: Page,
  htmlString: string
): Promise<string> => {
  return await page.evaluate((innerText: string): string => {
    const wrapper: HTMLDivElement = document.createElement("div");
    wrapper.innerHTML = innerText;

    return (wrapper.firstChild as HTMLElement).innerHTML;
  }, htmlString);
};

test("getCalendar", async ({ page }) => {
  // Go to econnect login page
  await page.goto("https://econnect.fourtitude.asia/Account/Login");

  // Login to the portal
  const loginContainer = await getLocator(page.getByPlaceholder("Login ID"));
  const passwordContainer = await getLocator(page.getByPlaceholder("Password"));

  await loginContainer[0].fill(globalVariables.USERID);
  await passwordContainer[0].fill(globalVariables.PASSWORD);

  const loginButton = page.getByRole("button", { name: "Login" });
  await loginButton.click();

  await page.waitForLoadState("networkidle");

  // After login go the leave application page
  await page.goto("https://econnect.fourtitude.asia/LMS/LeaveApplication");

  // Choose AL from the leave request dropdown
  const categoryDropdown = await getLocator(page.locator("#LeaveCategory"));
  categoryDropdown[0].selectOption("AL");

  // Wait for the calendar to appear
  const calendarContainer = await getLocator(
    page.locator(".zabuto_calendar"),
    60000
  );

  await page.pause();

  // Go to previous calendar month
  const tableHeaderButtons = await getLocator(
    calendarContainer[0].locator(".table > .calendar-month-header > th")
  );
  const previousMonthButton = tableHeaderButtons[0];
  await previousMonthButton.click();
  await fullLoadingDone(page);

  // isFull
  await getLocator(page.locator(".isFull"), 30000, true);

  // Getting all the td from the table for all the days in the month
  const allTd: Array<Locator> = [];
  const allDateRows = await getLocator(page.locator(".calendar-dow"));
  for (const row of allDateRows) {
    const rowTd = await getLocator(row.locator("td"));
    allTd.push(...rowTd);
  }

  // Getting the working days only (exclude the Sat and Sun)
  const allWorkingDays: Array<Locator> = [];
  for (const tds of allTd) {
    const dayTypeId = await tds.getAttribute("data-daytypeid"); // This is for the off day and rest day tooltip
    if (dayTypeId != "1" && dayTypeId != "2") {
      // 1 is rest day, 2 is off day
      allWorkingDays.push(tds);
    }
  }

  // Getting the days that can claim
  const daysCanClaim: Array<ClaimDay> = [];
  for (const work of allWorkingDays) {
    const result: ClaimDay = {};
    const dayTypeId = await work.getAttribute("data-daytypeid"); // This is for PH (id = 4)
    const eventTooltip = await work.getAttribute("event-tooltip"); // This is for the leave tooltip

    if (eventTooltip || dayTypeId == "4") {
      result.type = "leave";
    }

    if (eventTooltip) {
      // This is for leave
      result.desc = await getHtmlStringInnerHTML(page, eventTooltip);
    } else if (dayTypeId == "4") {
      // This is for PH
      const eventAtt = await work.getAttribute("event-scheduler");
      if (eventAtt) {
        result.desc = await getHtmlStringInnerHTML(page, eventAtt);
      }
    } else {
      result.type = "claim";
      result.desc = globalVariables.DESCRIPTION;
    }

    const innerText = await work.innerHTML();

    let dateContainer: Array<Locator> | null = null;
    if (innerText.includes("<span")) {
      dateContainer = await getLocator(work.locator("div > span"));
    } else if (innerText.includes("<div")) {
      dateContainer = await getLocator(work.locator("div"));
    }

    if (dateContainer) {
      result.day = Number(await dateContainer[0].innerHTML());
      daysCanClaim.push(result);
    }
  }

  const month = new Date().getMonth();
  const excelData: ExcelData = {
    month: globalVariables.ALL_MONTHS[month - 1 < 0 ? 11 : month - 1],
    days: daysCanClaim,
  };

  generateMealExcel(excelData);
});
