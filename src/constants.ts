import { EnumAction, EnumTargetType, EnumWeekday } from "./enums";

const WEEKDAYS = [
  EnumWeekday.MONDAY,
  EnumWeekday.TUESDAY,
  EnumWeekday.WEDNESDAY,
  EnumWeekday.THURSDAY,
  EnumWeekday.FRIDAY,
  EnumWeekday.SATURDAY,
  EnumWeekday.SUNDAY,
];

const ACTIONS = [EnumAction.ARCHIVE, EnumAction.DELETE, EnumAction.SPAM];

const TARGET_TYPES = [EnumTargetType.DOMAIN, EnumTargetType.EMAIL];

const LAST_UPDATE_PROPERTY = "LAST_UPDATE";

export { ACTIONS, LAST_UPDATE_PROPERTY, TARGET_TYPES, WEEKDAYS };
