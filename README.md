# Microsoft Teams Messaging Extension Middleware for Microsoft Bot Builder

[![npm version](https://badge.fury.io/js/botbuilder-teams-messagingextensions.svg)](https://badge.fury.io/js/botbuilder-teams-messagingextensions)

This middleware for [Bot Builder Framework](https://www.npmjs.com/package/botbuilder) is targeted for [Microsoft Teams](https://docs.microsoft.com/en-us/microsoftteams/platform/) based bots.

 | @master | @preview |
 :--------:|:---------:
 [![Build Status](https://travis-ci.org/wictorwilen/botbuilder-teams-messagingextensions.svg?branch=master)](https://travis-ci.org/wictorwilen/botbuilder-teams-messagingextensions)|[![Build Status](https://travis-ci.org/wictorwilen/botbuilder-teams-messagingextensions.svg?branch=preview)](https://travis-ci.org/wictorwilen/botbuilder-teams-messagingextensions)

## About

The Microsoft Teams [Messaging Extension](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/messaging-extensions-overview?view=msteams-client-js-latest) Middleware for Microsoft Bot Builder makes building bots for Microsoft Teams easier. By separating out the logic for Message Extensions from the implementation of the bot, you will make your code more readable and easier to debug and troubleshoot.

The middleware supports the following Message Extension features

* [Message extension queries](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/search-extensions): `composeExtension/query`
* [Message extension settings url](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/search-extensions#add-event-handlers): `composeExtension/querySettingUrl`
* [Message extension settings](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/search-extensions#add-event-handlers): `composeExtension/setting`
* [Message extension link unfurling](https://developer.microsoft.com/en-us/office/blogs/add-rich-previews-to-messages-using-link-unfurling/): `composeExtension/queryLink`
* [Message extension message actions](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/create-extensions): `composeExtension/submitAction`
* [Fetch task operations for message actions](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/create-extensions): `composeExtension/fetchTask`
* Adaptive Card `Action.Submit` actions: `composeExtension/onCardButtonClicked`
* [Message extension select](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/search-extensions): `composeExtension/selectItem`

## Usage

To implement a Messaging Extension handler create a class like this:

> NOTE: When combining this with the `botbuilder-teams` you should avoid using the  `invokeActivityHandler.onInvoke`, as it might
> invalidate your messaging extension results.

``` TypeScript
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";

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
            title: "Configuration",
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

To add the processor to the pipeline use code similar to this:

``` TypeScript
import { MessagingExtensionMiddleware } from "botbuilder-teams-messagingextensions";

const adapter = new BotFrameworkAdapter(botSettings);
adapter.user(new MessagingExtensionMiddleware("myCommandId", new MyMessageExtension()));
```

Where you should match the command id with the one in the Teams manifest file:

``` JSON
"composeExtensions": [{
    "botId": "12341234-1234-1234-123412341234",
    "canUpdateConfiguration": true,
    "commands": [{
        "id": "myCommandId",
        "title": "My Command",
        "description": "...",
        "initialRun": true,
        "parameters": [...]
    }]
}],
```

### Use message actions and task modules

To create an message action that shows a task module for your input define your message extension as follows in the manifest. The `fetchTask` property set to `true` indicates that we want to use a task module.

``` JSON
{
    "id": "createToDoMessageExtension",
    "title": "Create To-Do",
    "description": "Create a To-Do item",
    "context": ["message", "commandBox", "compose"],
    "fetchTask": true,
    "type": "action"
}
```

In the processor you need to implement the `onFetchTask` and `onSubmitAction` methods. You can either return a card using the `card` property or
use the `url` parameter to point to a web page.

``` TypeScript
public async onFetchTask(context: TurnContext, value: MessagingExtensionAction): Promise<MessagingExtensionResult | TaskModuleContinueResponse> {
    return Promise.resolve<ITaskModuleResult>({
        type: "continue",
            value: {
            title: "Task Module",
            card: CardFactory.adaptiveCard({
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                type: "AdaptiveCard",
                version: "1.0",
                body: [
                    {
                        type: "TextBlock",
                        text: "Please enter your e-mail"
                    },
                    {
                        type: "Input.Text",
                        id: "myEmail",
                        placeholder: "youremail@example.com",
                        style: "email"
                    },
                ],
                actions: [
                    {
                        type: "Action.Submit",
                        title: "OK",
                        data: { id: "unique-id" }
                    }
                ]
            })
        }
    });
}

// handle response in here
public async onSubmitAction(context: TurnContext, value: MessagingExtensionAction): Promise<MessagingExtensionResult> {
   const email = value.data.myEmail;
   const id = value.data.id;
   ...
}
```

## Contributors

* [Wictor Wilén](https://github.com/wictorwilen) - Original author and coordinator
* [Thomas White](https://github.com/tdwhite0)
* [Bill Bliss](https://github.com/billbliss)
* [greyseer256](https://github.com/greyseer256)
* [Kavin Singh](https://github.com/kavins14)

## License

Copyright (c) Wictor Wilén. All rights reserved.

Licensed under the MIT license.
