// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.

import { debug } from "debug";
import { Middleware, TurnContext } from "botbuilder";
import { ActivityTypesEx, MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder-teams";

// Initialize debug logging module
var log = debug("msteams");

/**
 * see https://raw.githubusercontent.com/OfficeDev/BotBuilder-MicrosoftTeams-node/
 * 2b1a9de550b7d724e38cbfad4ea96de7c4966900/botbuilder-teams-js/swagger/teamsAPI.json
 */
export interface ITaskModuleTaskInfo {
    title: string;
    height?: number | string | "small" | "medium" | "large";
    width?: number | string | "small" | "medium" | "large";
    url?: string;
    card?: any;
    fallbackUrl?: string;
    completionBotId?: string;
}

/**
 * see https://raw.githubusercontent.com/OfficeDev/BotBuilder-MicrosoftTeams-node/
 * 2b1a9de550b7d724e38cbfad4ea96de7c4966900/botbuilder-teams-js/swagger/teamsAPI.json
 */
export interface ITaskModuleResult {
    type: "message" | "continue";
    value: ITaskModuleTaskInfo;
}

/**
 * see https://raw.githubusercontent.com/OfficeDev/BotBuilder-MicrosoftTeams-node/
 * 2b1a9de550b7d724e38cbfad4ea96de7c4966900/botbuilder-teams-js/swagger/teamsAPI.json
 */
export interface IMessagingExtensionActionRequest {
    commandId: string;
    commandContext: "" | "message" | "compose" | "commandBox";
    context: {
        theme: string;
    };
    /**
     * `data` is sent back from an adaptive card, task module or static properties
     */
    data?: any;
    /**
     * `state` is sent back from a config/auth request
     */
    state?: any;
    messagePayload?: any;
    botMessagePreviewAction?: "edit" | "send";
}

/**
 * see https://raw.githubusercontent.com/OfficeDev/BotBuilder-MicrosoftTeams-node/
 * 2b1a9de550b7d724e38cbfad4ea96de7c4966900/botbuilder-teams-js/swagger/teamsAPI.json
 */
export interface IAppBasedLinkQuery {
    url: string;
}

// tslint:disable: max-line-length

/**
 * Defines the processor for the Messaging Extension Middleware
 */
export interface IMessagingExtensionMiddlewareProcessor {
    /**
     * Processes incoming queries (composeExtension/query)
     * @param context the turn context
     * @param value the value of the query
     * @returns {Promise<MessagingExtensionResult>}
     */
    onQuery?(context: TurnContext, value: MessagingExtensionQuery): Promise<MessagingExtensionResult>;
    /**
     * Process incoming requests for Messaging Extension settings (composeExtension/querySettingUrl)
     * @param context the turn context
     * @returns {Promise<{ title: string, value: string }}
     */
    onQuerySettingsUrl?(context: TurnContext): Promise<{ title: string, value: string }>;
    /**
     * Processes incoming setting updates (composeExtension/setting)
     * @param context the turn context
     * @returns {Promise<void>}
     */
    onSettings?(context: TurnContext): Promise<void>;
    /**
     * Processes incoming link queries (composeExtension/queryLink)
     * @param context the turn context
     * @param value the value of the query
     * @returns {Promise<MessagingExtensionResult>}
     */
    onQueryLink?(context: TurnContext, value: IAppBasedLinkQuery): Promise<MessagingExtensionResult>;
    /**
     * Processes incoming link actions (composeExtension/submitAction)
     * @param context the turn context
     * @param value the value of the query
     * @returns {Promise<MessagingExtensionResult>}
     */
    onSubmitAction?(context: TurnContext, value: IMessagingExtensionActionRequest): Promise<MessagingExtensionResult>;
    /**
     * Processes incoming fetch task actions (`composeExtension/fetchTask`)
     * @param context the turn context
     * @param value commandContext
     * @returns {Promise<MessagingExtensionResult | ITaskModuleResult>} Promise object is either a `MessagingExtensionResult` for `conf` or `auth` or a `ITaskModuleResult` for `message` or `continue`
     */
    onFetchTask?(context: TurnContext, value: IMessagingExtensionActionRequest): Promise<MessagingExtensionResult | ITaskModuleResult>;
    /**
     * Handles Action.Submit from adaptive cards
     *
     * Note: this is experimental and it does not filter on the commandId which means that if there are
     * multiple registered message extension processors all will recieve this command. You should ensure to
     * add a specific identifier to your adaptivecard.
     * @param context the turn context
     * @param value the card data
     * @returns {Promise<void>}
     */
    onCardButtonClicked?(context: TurnContext, value: any): Promise<void>;

    /**
     * Handles when an item is selected from the result list
     *
     * Note: this is experimental and it does not filter on the commandId which means that if there are
     * multiple registered message extension processors all will recieve this command. You should ensure to
     * add a specific identifier to your invoke action.
     * @param context the turn context
     * @param value object passed in to invoke action
     * @returns {Promise<MessagingExtensionResult>}
     */
    onSelectItem?(context: TurnContext, value: any): Promise<MessagingExtensionResult>;
}

/**
 * A Messaging Extension Middleware for Microsoft Teams
 */
export class MessagingExtensionMiddleware implements Middleware {

    /**
     * Default constructor
     * @param commandId The commandId of the Messaging Extension to process,
     *                  or `undefined` to process all incoming requests
     * @param processor The processor
     */
    public constructor(
        private commandId: string | undefined,
        private processor: IMessagingExtensionMiddlewareProcessor) {

    }
    
    /**
     * Bot Framework `onTurn` method
     * @param context the turn context
     * @param next the next function
     */
    public async onTurn(context: TurnContext, next: () => Promise<void>): Promise<void> {
        log(`Activity received - activity.name: ${context.activity.name}`);
        if (this.commandId !== undefined) {
            log(`  commandId: ${context.activity.value.commandId}`);
            log(`  parameters: ${JSON.stringify(context.activity.value.parameters)}`);
        } else {
            log(`  activity.value: ${JSON.stringify(context.activity.value)}`);
        }
        if ((this.commandId !== undefined) && 
            (context.activity.value.commandId.toLowerCase() !== this.processor.constructor.name.toLowerCase())) {
            // command id from manifest.json not bound to a middleware handler
            log(`ERROR: the command id defined in manifest.json [${context.activity.value.commandId}] does not match the middleware handler [${this.processor.constructor.name}] and will not be handled.`);
            log(`  Your messaging extension bot will return an HTTP 501 "Not Implemented" error.`);
            log(`  If this app was generated by the Teams Yeoman Generator, ${this.processor.constructor.name} is implemented in ${this.processor.constructor.name}.ts.`);
        }
        switch (context.activity.name) {
            case "composeExtension/query":
                if ((this.commandId === context.activity.value.commandId || this.commandId === undefined) &&
                    this.processor.onQuery) {
                    try {
                        const result = await this.processor.onQuery(context, context.activity.value);
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: {
                                    composeExtension: result,
                                },
                                status: 200,
                            },
                        });
                    } catch (err) {
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: err,
                                status: 500,
                            },
                        });
                    }
                    return;
                }
                break;
            case "composeExtension/querySettingUrl":
                if ((this.commandId === context.activity.value.commandId || this.commandId === undefined) &&
                    this.processor.onQuerySettingsUrl) {
                    try {
                        const result = await this.processor.onQuerySettingsUrl(context);
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: {
                                    composeExtension: {
                                        suggestedActions: {
                                            actions: [{
                                                type: "openApp",
                                                ...result,
                                            }],
                                        },
                                        type: "config",
                                    },
                                },
                                status: 200,
                            },
                        });
                    } catch (err) {
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: err,
                                status: 500,
                            },
                        });
                    }
                    return;
                }
                break;
            case "composeExtension/setting":
                if ((this.commandId === context.activity.value.commandId || this.commandId === undefined) &&
                    this.processor.onSettings) {
                    try {
                        await this.processor.onSettings(context);
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                status: 200,
                            },
                        });
                    } catch (err) {
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: err,
                                status: 500,
                            },
                        });
                    }
                    return;
                }
                break;
            case "composeExtension/queryLink":
                if (this.processor.onQueryLink) {
                    try {
                        const result = await this.processor.onQueryLink(context, context.activity.value);
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: {
                                    composeExtension: result,
                                },
                                status: 200,
                            },
                        });
                    } catch (err) {
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: err,
                                status: 500,
                            },
                        });
                    }
                    return;
                    // we're doing a return here and not next() so we're not colliding with
                    // any botbuilder-teams invoke things. This however will also invalidate the use
                    // of multiple message extensions using queryLink - only the first one will be triggered
                }
                break;
            case "composeExtension/submitAction":
                if ((this.commandId === context.activity.value.commandId || this.commandId === undefined) &&
                    this.processor.onSubmitAction) {
                    try {
                        const result = await this.processor.onSubmitAction(context, context.activity.value);
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: {
                                    composeExtension: result,
                                },
                                status: 200,
                            },
                        });
                    } catch (err) {
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: err,
                                status: 500,
                            },
                        });
                    }
                    return;
                }
                break;
            case "composeExtension/fetchTask":
                if ((this.commandId === context.activity.value.commandId || this.commandId === undefined) &&
                    this.processor.onFetchTask) {
                    try {
                        const result = await this.processor.onFetchTask(context, context.activity.value);
                        const body = result.type === "continue" || result.type === "message" ?
                            { task: result } :
                            { composeExtension: result };
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body,
                                status: 200,
                            },
                        });
                    } catch (err) {
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: err,
                                status: 500,
                            },
                        });
                    }
                    return;
                }
                break;
            case "composeExtension/onCardButtonClicked":
                if (this.processor.onCardButtonClicked) {
                    try {
                        await this.processor.onCardButtonClicked(context, context.activity.value);
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                status: 200,
                            },
                        });
                    } catch (err) {
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: err,
                                status: 500,
                            },
                        });
                    }
                }
                break;
            case "composeExtension/selectItem":
                if (this.processor.onSelectItem) {
                    try {
                        const result = await this.processor.onSelectItem(context, context.activity.value);
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: {
                                    composeExtension: result,
                                },
                                status: 200,
                            },
                        });
                        return;
                        // we're doing a return here and not next() so we're not colliding with
                        // any botbuilder-teams invoke things. This however will also invalidate the use
                        // of multiple message extensions using selectItem - only the first one will be triggered
                    } catch (err) {
                        context.sendActivity({
                            type: ActivityTypesEx.InvokeResponse,
                            value: {
                                body: err,
                                status: 200,
                            },
                        });
                    }
                }
                break;
            default:
                // nop
                break;
        }
        return next();
    }
}
