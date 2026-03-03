declare interface INotificationsCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'NotificationsCommandSetStrings' {
  const strings: INotificationsCommandSetStrings;
  export = strings;
}
