enum EnumAction {
  ARCHIVE = "Archive",
  DELETE = "Delete",
  SPAM = "Spam",
}

enum EnumColor {
  GREEN = "#34A853",
  RED = "#EA4335",
}

enum EnumSheet {
  INBOX = "Inbox",
  ACTIONS = "Actions",
  LOG = "Log",
  EMAIL_PIVOT = "#Email",
  DOMAIN_PIVOT = "#Domain",
  HOURS_PIVOT = "#Hours",
}

enum EnumInboxColumn {
  THREAD_ID = 0,
  MAIL_ID = 1,
  EMAIL = 2,
  EMAIL_DOMAIN = 3,
  DATE = 4,
  SUBJECT = 5,
  WEEKDAY = 6,
  HOUR = 7,
}

enum EnumTargetType {
  DOMAIN = "Domain",
  EMAIL = "Email",
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
