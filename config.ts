enum Sheet {
  INBOX = "Inbox",
  ACTIONS = "Actions",
  EMAIL_PIVOT = "#Email",
  DOMAIN_PIVOT = "#Domain",
  HOURS_PIVOT = "#Hours",
}

enum Weekday {
  MONDAY = "Mon",
  TUESDAY = "Tue",
  WEDNESDAY = "Wed",
  THURSDAY = "Thu",
  FRIDAY = "Fri",
  SATURDAY = "Sat",
  SUNDAY = "Sun",
}

const WEEKDAYS = [
  Weekday.MONDAY,
  Weekday.TUESDAY,
  Weekday.WEDNESDAY,
  Weekday.THURSDAY,
  Weekday.FRIDAY,
  Weekday.SATURDAY,
  Weekday.SUNDAY,
];

enum Color {
  GREEN = "#34A853",
  RED = "#EA4335"
}

enum Action {
  ARCHIVE = "Archive",
  DELETE = "Delete",
  SPAM = "Spam",
}

const ACTIONS = [
  Action.ARCHIVE,
  Action.DELETE,
  Action.SPAM,
];

enum TargetType {
  DOMAIN = "Domain",
  EMAIL = "Email",
}

const TARGET_TYPES = [
  TargetType.DOMAIN,
  TargetType.EMAIL,
];