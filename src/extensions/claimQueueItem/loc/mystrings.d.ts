declare interface IClaimQueueItemCommandSetStrings {
  CommandLabel: string;
  SuccessMessage: string;
  AlreadyTakenMessage: string;
  UnexpectedErrorMessage: string;
}

declare module 'ClaimQueueItemCommandSetStrings' {
  const strings: IClaimQueueItemCommandSetStrings;
  export = strings;
}
