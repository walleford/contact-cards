import { WebPartContext } from '@microsoft/sp-webpart-base'

export interface ICallingCardsProps {
    description: string;
    spfxContext: WebPartContext;
    CallingCards: any[];
    Layout: string;
}
