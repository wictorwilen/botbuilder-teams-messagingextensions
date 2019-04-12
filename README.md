# Microsoft Teams Messaging Extension Middleware for Microsoft Bot Builder 

[![npm version](https://badge.fury.io/js/botbuilder-teams-messagingextensions.svg)](https://badge.fury.io/js/botbuilder-teams-messagingextensions)

This middleware for [Bot Builder Framework](https://www.npmjs.com/package/botbuilder) is targeted for [Microsoft Teams](https://docs.microsoft.com/en-us/microsoftteams/platform/) based bots.

 | @master | @preview |
 :--------:|:---------:
 [![Build Status](https://travis-ci.org/wictorwilen/botbuilder-teams-messagingextensions.svg?branch=master)](https://travis-ci.org/wictorwilen/botbuilder-teams-messagingextensions)|[![Build Status](https://travis-ci.org/wictorwilen/botbuilder-teams-messagingextensions.svg?branch=preview)](https://travis-ci.org/wictorwilen/botbuilder-teams-messagingextensions)

TODO: 

## Usage

To implement a Messaging Extension handler create a class like this:

``` TypeScript
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";

@PreventIframe("/todoTeamsMessageExtension/config.html")
export default class MyMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQuery(query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
        const card = CardFactory.heroCard("Test", "Test", ["https://picsum.photos/200/200"]);

        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            return Promise.resolve({
                type: "result",
                attachmentLayout: "grid",
                attachments: [
                    card
                ]
            });
        } else {
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: [
                    card
                ]
            });
        }
    }

    public async onQuerySettingsUrl(): Promise<{ title: string, value: string }> {
        return Promise.resolve({
            title: "Todo Teams Configuration",
            value: "https://my-service-com/config.html"
        });
    }

    public async onSettingsUpdate(context: TurnContext): Promise<void> {
        const setting = context.activity.value.state;
        // Save the setting
        return Promise.resolve();
    }
}
``` 

## License

Copyright (c) Wictor Wilén. All rights reserved.

Licensed under the MIT license.