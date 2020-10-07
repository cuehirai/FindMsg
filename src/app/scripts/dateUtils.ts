// DOCS:  https://date-fns.org/v2.15.0/docs/format

export { default as differenceInMilliseconds } from "date-fns/differenceInMilliseconds";
export { default as endOfDay } from "date-fns/endOfDay";
export { default as endOfToday } from "date-fns/endOfToday";
export { default as format } from "date-fns/format";
export { default as formatISO } from "date-fns/formatISO";
export { default as isAfter } from "date-fns/isAfter";
export { default as isBefore } from "date-fns/isBefore";
export { default as isDate } from "date-fns/isDate";
export { default as isValid } from "date-fns/isValid";
export { default as max } from "date-fns/max";
export { default as parseISO } from "date-fns/parseISO";
export { default as startOfDay } from "date-fns/startOfDay";
export { default as startOfToday } from "date-fns/startOfToday";
export { default as subDays } from "date-fns/subDays";
export { default as subHours } from "date-fns/subHours";
export { default as subMinutes } from "date-fns/subMinutes";
export { default as subMonths } from "date-fns/subMonths";
export { default as subWeeks } from "date-fns/subWeeks";
export { default as subYears } from "date-fns/subYears";
//export { default as  } from "date-fns/";

export const invalidDate = (): Date => new Date(NaN);
export const now = (): Date => new Date();

import { default as isValid } from "date-fns/isValid";

/**
 * Convert a date to a timestamp number, map invalid date to -1
 * @param m
 */
export function dateToNumber(m: Date): number {
    return (isValid(m) ? m.getTime() : -1);
}

/**
 * Convert a timestamp number to a date instance, treating negative numbers as invalid
 * @param n
 */
export function numberToDate(n: number): Date {
    return n < 0 ? invalidDate() : new Date(n);
}
