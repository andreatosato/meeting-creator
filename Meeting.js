import React, { Component, Suspense } from 'react';
import { Button, Flex } from '@fluentui/react-northstar';
import * as Msal from 'msal';
import { Client, MicrosoftGraph } from "@microsoft/microsoft-graph-client";


class Meeting extends Component {
  
  constructor(props) {
    super(props);
    this.msalConfig = {
      auth: {
        clientId: "your_client_id", // Client Id of the registered application
        redirectUri: "your_redirect_uri",
      },
    };
    this.graphScopes = ["user.read", "mail.send"]; // An array of graph scopes
    this.msalApplication = new Msal.UserAgentApplication(this.msalConfig);
    this.options = new MicrosoftGraph.MSALAuthenticationProviderOptions(this.graphScopes);
    this.authProvider = new MicrosoftGraph.ImplicitMSALAuthenticationProvider(msalApplication, this.options);
    this.Client = MicrosoftGraph.Client;
    this.client = Client.initWithMiddleware(options);
    try {
      let userDetails = await client.api("/me").get();
      console.log(userDetails);
    } catch (error) {
      throw error;
    }
  }

  createMeeting(){
    console.log("create meeting");
  }

  render() {
      return (
          <div className="Meeting" style={{ height: '100vh' }}>
            <Flex gap="gap.small">
            <Button content="Create Meeting" primary onClick={this.createMeeting} />
          </Flex>
          </div>
      );
  }
}

export default Meeting;