import { expect, Locator, Page } from "@playwright/test";

const maxTry: number = 5;
let numOfTries: number = 0;
let countInterval: NodeJS.Timeout;
let forceResolve: any = null;

const allMonths = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

interface constants {
  USERID: string;
  PASSWORD: string;
  FULLNAME: string;
  SHORTNAME: string;
  SUPERVISOR: string;
  POSITION: string;
  COMPANY: string;
  ALL_MONTHS: Array<string>;
  DESCRIPTION: string;
  LOCATION: string;
  START_TIME: string;
  END_TIME: string;
  AMOUNT: number;
  EXCEL_FILENAME: string;
  SAVE_LOCATION: string;
}

export const globalVariables: constants = {
  USERID: "FM177",
  PASSWORD: "Alyafatysa_3113",

  ALL_MONTHS: [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ],

  // For excel headers
  FULLNAME: "Muhammad Azizi bin Abdul Aziz", // Your full name to save in the excel header
  SHORTNAME: "Azizi", // Short name for the excel file name
  SUPERVISOR: "Chong Hsu-Cherng",
  POSITION: "FRONT-END DEVELOPER",
  COMPANY: "FUEL MEDIA SDN BHD",
  DESCRIPTION: "Meal Allowance",
  LOCATION: "AIS/BS",
  START_TIME: "9:00 AM",
  END_TIME: "6:00 PM",
  AMOUNT: 10,
  EXCEL_FILENAME: "027-HRA-FRM-MRF-Rev 2 (Meal Reimbursement Form)_FMSB",
  SAVE_LOCATION:
    "C:\\Users\\FUEL-Muhammad.Azizi\\Desktop\\New folder Testing\\New folder",
};

const checkCountInterval = async (locator: Locator) => {
  const count = await locator.count();

  if (count == 0) {
    numOfTries++;
  } else {
    forceResolve(count);
    clearInterval(countInterval);
  }

  if (numOfTries == maxTry) {
    clearInterval(countInterval);
    throw "Locator not found after force";
  }
};

const forceAppear = async (locator: Locator): Promise<number> => {
  numOfTries = 0;

  const prom = new Promise<number>((res) => (forceResolve = res));
  countInterval = setInterval(() => checkCountInterval(locator), 1000);

  return prom;
};

export const getLocator = async (
  locator: Locator,
  timeout: number = 5000,
  force: boolean = false
): Promise<Array<Locator>> => {
  const arrayResult: Array<Locator> = [];
  let count = await locator.count();

  if (count == 0) {
    if (force) {
      count = await forceAppear(locator);
    } else {
      throw `Locator is not found`;
    }
  }

  for (let i = 0; i < count; i++) {
    await expect(locator.nth(i)).toBeVisible({ timeout: timeout });
    arrayResult.push(locator.nth(i));
  }

  return arrayResult;
};

export const fullLoadingDone = async (
  page: Page,
  loadingId: string = "#loadingIconInet"
): Promise<void> => {
  const loadingContainer = page.locator("#loadingIconInet");
  await expect(loadingContainer).toBeHidden();
};

export const getDayWithSuffix = (day: number): string => {
  if (day >= 11 && day <= 13) {
    return `${day}th`;
  }

  let suffix = "";
  if (day % 10 == 1) {
    suffix = "st";
  } else if (day % 10 == 2) {
    suffix = "nd";
  } else if (day % 10 == 3) {
    suffix = "rd";
  } else {
    suffix = "th";
  }
  return `${day}${suffix}`;
};
