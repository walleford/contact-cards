import * as React from 'react';
import styles from './CallingCards.module.scss';
import { ICallingCardsProps } from './ICallingCardsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ICallingCardsState } from './ICallingCardsState';
import { spfi, SPFx } from '@pnp/sp';


export default class CallingCards extends React.Component<ICallingCardsProps, ICallingCardsState> {

    private sp = spfi().using(SPFx(this.context));

    constructor(props: ICallingCardsProps, state: ICallingCardsState) {
        super(props);
        this.state = {
            description: this.props.description,
            CallingCards: this.props.CallingCards || [],
            Layout: this.props.Layout,
        };
        this._getCallingCards;
    }

    private _getCallingCards() {
        if (this.state.CallingCards.length > 0) {
            this.setState({ CallingCards: this.state.CallingCards });
        }
    }

    public componentDidUpdate(prevProps: ICallingCardsProps, prevState: ICallingCardsState): void {
        // If properties have changed bind it and update webpart
        if (this.props.CallingCards !== prevProps.CallingCards && this.props.CallingCards.length !== 0) {
            this.setState({ CallingCards: this.props.CallingCards });
        }
    }

    public updateRenderingStyle(): string {
        let style;
        if (this.props.Layout === 'vertical') {
            style = styles.block
            return style
        } else if (this.props.Layout === 'horizontal') {
            style = styles.grid
            return style
        }
    }

    public renderBioLink(el): string {
        let link;
        let noLink;
        if (el.bioLink === '') {
            link = <div className={`${styles.nameStyles}`} key={el}>{el.Name}</div>;
            return link;
        } else {
            link = <div className={`${styles.bioStyles}`} key={el} onClick={this.openLink(el)}>{el.Name}</div>
            return link
        }
    }

    private openLink(el): any {
        window.open(el)
    }
        
    public render(): React.ReactElement<ICallingCardsProps> {

        let arr = this.props.CallingCards || [];
        console.log(this.props.CallingCards);

        
        if (this.props.CallingCards && this.props.CallingCards.length > 0) {
            var contactsPicture = arr.map(el =>
                <div className={`${styles.tile}`}>            
                    <img className={`${styles.leadershipImage}`} key={el} src={el.filePicker.fileAbsoluteUrl} />
                    <div className={`${styles.textContainer}`}>
                        <a href={el.bioLink ? el.bioLink : null}>
                            <div className={`${el.bioLink ? styles.bioStyles : styles.nameStyles}`} key={el}>{el.Name}</div>
                        </a>
                        <div className={`${styles.textStyles}`} key={el}>{el.Branch}</div>
                        <div className={`${styles.textStyles}`} key={el}>{el.Position}</div>
                        <div className={`${styles.textStyles}`} key={el}>{el.PhoneNumber}</div>
                        <div className={`${styles.textStyles}`} key={el}>{el.dsn}</div>
                        <div className={`${styles.textStyles}`} key={el}>{el.duty}</div>
                        <span><span className={`${styles.textStyles}`} key={el}>Email: </span><a className={`${styles.emailStyles}`} href={`mailto:${el.Email}`} target="_top" key={el}>{el.Email}</a></span>
                    </div>
                </div>
            )
        } else {
            return (
                <div className={`${styles.welcome}`}>Use property pane to create new Contact Cards!</div>
            )
        }

        return (
            <body>
                <div className={this.updateRenderingStyle()}>
                    {contactsPicture}
                </div>
            </body>
        );
    }
}