enum EnumSheet {
  INBOX = "Inbox",
  ACTIONS = "Actions",
  EMAIL_PIVOT = "#Email",
  DOMAIN_PIVOT = "#Domain",
  HOURS_PIVOT = "#Hours",
}

enum EnumWeekday {
  MONDAY = "Mon",
  TUESDAY = "Tue",
  WEDNESDAY = "Wed",
  THURSDAY = "Thu",
  FRIDAY = "Fri",
  SATURDAY = "Sat",
  SUNDAY = "Sun",
}

const WEEKDAYS = [
  EnumWeekday.MONDAY,
  EnumWeekday.TUESDAY,
  EnumWeekday.WEDNESDAY,
  EnumWeekday.THURSDAY,
  EnumWeekday.FRIDAY,
  EnumWeekday.SATURDAY,
  EnumWeekday.SUNDAY,
];

enum EnumColor {
  GREEN = "#34A853",
  RED = "#EA4335"
}

enum EnumAction {
  ARCHIVE = "Archive",
  DELETE = "Delete",
  SPAM = "Spam",
}

const ACTIONS = [
  EnumAction.ARCHIVE,
  EnumAction.DELETE,
  EnumAction.SPAM,
];

enum EnumTargetType {
  DOMAIN = "Domain",
  EMAIL = "Email",
}

const TARGET_TYPES = [
  EnumTargetType.DOMAIN,
  EnumTargetType.EMAIL,
];

const LAST_UPDATE_PROPERTY = "LAST_UPDATE";

export {
  EnumSheet,
  EnumWeekday,
  EnumColor,
  EnumAction,
  EnumTargetType,
  ACTIONS,
  LAST_UPDATE_PROPERTY,
  TARGET_TYPES,
  WEEKDAYS,
};
