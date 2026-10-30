declare interface IEmployeeOperationsCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'EmployeeOperationsCommandSetStrings' {
  const strings: IEmployeeOperationsCommandSetStrings;
  export = strings;
}
