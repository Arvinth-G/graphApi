import * as React from 'react';
import { ITestWebpartProps } from './ITestWebpartProps';
import { MSGraphClient } from '@microsoft/sp-http';

export default class TestWebpart extends React.Component<ITestWebpartProps, {}> {

  constructor(props) {
    super(props)
    this.state = { displayName: '' }
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('/users/ArvinthGanesan@TestTrichy2.onmicrosoft.com')
          .get((error, response: any, rawResponse?: any) => {
            console.log(JSON.stringify(response));
            this.setState({
              displayName: response['displayName']
            })
          })
      });
  }

  public render(): React.ReactElement<ITestWebpartProps> {
    return (
      <div>
        <h4>{this.state['displayName']}</h4>
      </div>
    );
  }
}
