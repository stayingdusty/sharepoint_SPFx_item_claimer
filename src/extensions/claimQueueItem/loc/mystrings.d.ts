declare interface IClaimQueueItemCommandSetStrings {
  CommandLabel: string;
  ClaimingLabel: string;
  ClaimedLabel: string;
  SuccessMessage: string;
  AlreadyTakenMessage: string;
  UnexpectedErrorMessage: string;
}

declare module 'ClaimQueueItemCommandSetStrings' {
  const strings: IClaimQueueItemCommandSetStrings;
  export = strings;
}
