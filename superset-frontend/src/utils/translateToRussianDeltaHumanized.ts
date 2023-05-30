import { t } from '@superset-ui/core';

function translateToRussianDeltaHumanized(timeString: string): string | null {
  try {
    const numberString = timeString.match(/\d+/)?.[0]; // match the number part
    if (!numberString) return t(timeString);
    const number = Number(numberString); // convert string to number
    const unit = timeString.match(/[a-zA-Z]+/)?.[0]; // match the unit part

    const rtf = new Intl.RelativeTimeFormat('ru', { numeric: 'auto' });

    // map English units to their Russian equivalents
    const unitMap: { [key: string]: Intl.RelativeTimeFormatUnit } = {
      second: 'second',
      minute: 'minute',
      hour: 'hour',
      day: 'day',
      week: 'week',
      month: 'month',
      year: 'year',
      seconds: 'second',
      minutes: 'minute',
      hours: 'hour',
      days: 'day',
      weeks: 'week',
      months: 'month',
      years: 'year',
    };

    const russianUnit = unitMap[unit!];

    if (!russianUnit) {
      return t(timeString);
    }

    // use negative number to represent "ago"
    return rtf.format(-number, russianUnit);
  } catch (error) {
    return null;
  }
}

export default translateToRussianDeltaHumanized;
