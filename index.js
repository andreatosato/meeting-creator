// https://github.com/microsoftgraph/msgraph-sdk-javascript
// https://docs.microsoft.com/en-us/javascript/api/overview/msteams-client?view=msteams-client-js-latest
// https://fluentsite.z22.web.core.windows.net/theming-examples

import React, { Component } from 'react';
import { render } from 'react-dom';
import Hello from './Hello';
import './style.css';
import { Provider, themes } from '@fluentui/react-northstar'
import * as microsoftTeams from "@microsoft/teams-js";
import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";
import Meeting from './Meeting';

class App extends Component {
  constructor() {
    super();
    this.state = {
      name: 'React',
      theme: themes.teams
    };
  }

  componentDidMount() {
    microsoftTeams.initialize();
    microsoftTeams.appInitialization.notifyAppLoaded();

    microsoftTeams.registerOnThemeChangeHandler(theme => this.switchTheme(theme));
    microsoftTeams.getContext(async context => {
        this.switchTheme(context.theme);
        console.log(JSON.stringify(context));
        localStorage.setItem('tid', context.tid);
        localStorage.setItem('upn', context.upn);
    });
  }
  switchTheme(theme){
    switch (theme) {
      case 'default':
        this.setState({ theme: themes.teams});
        break;
      case 'dark':
        this.setState({ theme: themes.teamsDark});
        break;
      case 'contrast':
        this.setState({ theme: themes.teamsHighContrast});
        break;
      default:
        this.setState({ theme: themes.teams});
        break;
    }
    this.forceUpdate();
  }

  render() {
    return (
      <Provider theme={this.state.theme}>
          <Meeting /> 
      </Provider>
    );
  }
}

render(<App />, document.getElementById('root'));
