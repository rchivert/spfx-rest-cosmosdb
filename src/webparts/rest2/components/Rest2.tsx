import * as React from 'react';
import styles from './Rest2.module.scss';
import { IRest2Props } from './IRest2Props';
import { escape } from '@microsoft/sp-lodash-subset';

import { autobind } from '@uifabric/utilities';

export default class Rest2 extends React.Component<IRest2Props, {}> {

  private myPromise : Promise<any>;

  @autobind
  public componentDidMount(): void {
    //
    //

    // this.myPromise = this.setAndgetData()
    // .then (x =>  {
    //   //  set state
    //   console.log("set state here...");
    //   });
    }


  public render(): React.ReactElement<IRest2Props> {
    return (
      <div className={ styles.rest2 }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
