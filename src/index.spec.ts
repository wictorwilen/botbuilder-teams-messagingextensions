import { TurnContext } from "botbuilder-core";
import { stub } from "jest-auto-stub";

import * as exp from "./index";

const next = jest.fn().mockResolvedValue(undefined);
const sendActivity = jest.fn().mockResolvedValue(undefined);

describe("index", () => {
    let processor: exp.IMessagingExtensionMiddlewareProcessor;
    beforeEach(() => {
        processor = stub<exp.IMessagingExtensionMiddlewareProcessor>();
        jest.resetAllMocks();
    });

    it("Should export MessagingExtensionMiddleware", () => {
        expect(exp.MessagingExtensionMiddleware).toBeDefined();
    });

    it("Should successfully create the MessagingExtensionMiddleware", () => {
        const mw = new exp.MessagingExtensionMiddleware("command", processor);
        expect(mw).toBeDefined();
    });

    it("Should successfully call OnTurn and pass through", async () => {
        const mw = new exp.MessagingExtensionMiddleware("command", processor);
        const result = await mw.onTurn(stub<TurnContext>(), next);
        expect(result).toBe(undefined);
        expect(next).toBeCalled();
    });

    it("Should successfully call OnTurn and pass through, without activity", async () => {
        const mw = new exp.MessagingExtensionMiddleware("command", processor);
        const result = await mw.onTurn(stub<TurnContext>({ activity: undefined }), next);
        expect(result).toBe(undefined);
        expect(next).toBeCalled();
    });

    it("Should successfully call OnTurn and pass through, without activity name", async () => {
        const mw = new exp.MessagingExtensionMiddleware("command", processor);
        const result = await mw.onTurn(stub<TurnContext>({ activity: { name: undefined } }), next);
        expect(result).toBe(undefined);
        expect(next).toBeCalled();
    });

    describe("onActionExecute", () => {
        it("Should not call onActionExecute", async () => {
            processor.onActionExecute = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({ sendActivity }), next);
            expect(result).toBe(undefined);
            expect(processor.onActionExecute).toBeCalledTimes(0);
            expect(sendActivity).not.toBeCalled();
            expect(next).toBeCalled();
        });

        it("Should call onActionExecute", async () => {
            processor.onActionExecute = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: { name: "adaptiveCard/action" },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(sendActivity).toBeCalledTimes(1);
            expect(processor.onActionExecute).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should handle onActionExecute error", async () => {
            processor.onActionExecute = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: { name: "adaptiveCard/action" },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onActionExecute).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should call next if missing onActionExecute", async () => {
            processor.onActionExecute = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: { name: "adaptiveCard/action" },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(sendActivity).not.toBeCalled();
            expect(next).toBeCalled();
        });
    });

    describe("onQuery", () => {
        it("Should call onQuery", async () => {
            processor.onQuery = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/query",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onQuery).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should handle onQuery error", async () => {
            processor.onQuery = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/query",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onQuery).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should not call onQuery - invalid command id", async () => {
            processor.onQuery = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/query",
                    value: {
                        commandId: "wrong"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onQuery).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });

        it("Should not call onQuery - missing onQuery method", async () => {
            processor.onQuery = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/query",
                    value: {
                        commandId: "command"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(next).toBeCalled();
        });

        it("Should not call onQuery - not correct activity name", async () => {
            processor.onQuery = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onQuery).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });

    describe("onQuerySettingsUrl", () => {
        it("Should call onQuerySettingsUrl", async () => {
            processor.onQuerySettingsUrl = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/querySettingUrl",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onQuerySettingsUrl).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should handle onQuerySettingsUrl error", async () => {
            processor.onQuerySettingsUrl = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/querySettingUrl",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onQuerySettingsUrl).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should not call onQuerySettingsUrl - invalid command id", async () => {
            processor.onQuerySettingsUrl = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/querySettingUrl",
                    value: {
                        commandId: "wrong"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onQuerySettingsUrl).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });

        it("Should not call onQuerySettingsUrl - missing onQuery method", async () => {
            processor.onQuerySettingsUrl = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/querySettingUrl",
                    value: {
                        commandId: "command"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(next).toBeCalled();
        });

        it("Should not call onQuerySettingsUrl - not correct activity name", async () => {
            processor.onQuerySettingsUrl = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onQuerySettingsUrl).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });

    describe("onSettings", () => {
        it("Should call onSettings", async () => {
            processor.onSettings = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/setting",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onSettings).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should handle onSettings error", async () => {
            processor.onSettings = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/setting",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onSettings).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should not call onSettings - invalid command id", async () => {
            processor.onSettings = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/setting",
                    value: {
                        commandId: "wrong"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onSettings).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });

        it("Should not call onSettings - missing onQuery method", async () => {
            processor.onSettings = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/setting",
                    value: {
                        commandId: "command"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(next).toBeCalled();
        });

        it("Should not call onSettings - not correct activity name", async () => {
            processor.onSettings = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onSettings).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });

    describe("onQueryLink", () => {
        it("Should call onQueryLink", async () => {
            processor.onQueryLink = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/queryLink",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onQueryLink).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should handle onQueryLink error", async () => {
            processor.onQueryLink = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/queryLink",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onQueryLink).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should call onQueryLink - with invalid command id", async () => {
            processor.onQueryLink = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/queryLink",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onQueryLink).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should not call onQueryLink - missing onQuery method", async () => {
            processor.onQueryLink = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/queryLink",
                    value: {
                        commandId: "command"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(next).toBeCalled();
        });

        it("Should not call onQueryLink - not correct activity name", async () => {
            processor.onQueryLink = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onQueryLink).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });

    describe("onSelectItem", () => {
        it("Should call onSelectItem", async () => {
            processor.onSelectItem = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/selectItem",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onSelectItem).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should handle onSelectItem error", async () => {
            processor.onSelectItem = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/selectItem",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onSelectItem).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should call onSelectItem - with invalid command id", async () => {
            processor.onSelectItem = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/selectItem",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onSelectItem).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should not call onSelectItem - missing onQuery method", async () => {
            processor.onSelectItem = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/selectItem",
                    value: {
                        commandId: "command"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(next).toBeCalled();
        });

        it("Should not call onSelectItem - not correct activity name", async () => {
            processor.onSelectItem = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onSelectItem).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });

    describe("onCardButtonClicked", () => {
        it("Should call onCardButtonClicked", async () => {
            processor.onCardButtonClicked = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/onCardButtonClicked",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onCardButtonClicked).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).toBeCalled();
        });

        it("Should handle onCardButtonClicked error", async () => {
            processor.onCardButtonClicked = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/onCardButtonClicked",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onCardButtonClicked).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).toBeCalled();
        });

        it("Should call onCardButtonClicked - with invalid command id", async () => {
            processor.onCardButtonClicked = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/onCardButtonClicked",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onCardButtonClicked).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).toBeCalled();
        });

        it("Should not call onCardButtonClicked - missing onQuery method", async () => {
            processor.onCardButtonClicked = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/onCardButtonClicked",
                    value: {
                        commandId: "command"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(next).toBeCalled();
        });

        it("Should not call onCardButtonClicked - not correct activity name", async () => {
            processor.onCardButtonClicked = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onCardButtonClicked).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });

    describe("onFetchTask/fetchTask", () => {
        it("Should call onFetchTask", async () => {
            processor.onFetchTask = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/fetchTask",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should handle onFetchTask error", async () => {
            processor.onFetchTask = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/fetchTask",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should call onFetchTask - with invalid command id", async () => {
            processor.onFetchTask = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/fetchTask",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should not call onFetchTask - missing onQuery method", async () => {
            processor.onFetchTask = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/fetchTask",
                    value: {
                        commandId: "command"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(next).toBeCalled();
        });

        it("Should not call onFetchTask - not correct activity name", async () => {
            processor.onFetchTask = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });

    describe("onFetchTask/fetch", () => {
        it("Should call onFetchTask", async () => {
            processor.onFetchTask = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "task/fetch",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should call onFetchTask, without a commandId", async () => {
            processor.onFetchTask = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware(undefined, processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "task/fetch",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({
                type: "invokeResponse",
                value: { body: { composeExtension: {} }, status: 200 }
            });
            expect(next).not.toBeCalled();
        });

        it("Should call onFetchTask, returning a continue message", async () => {
            processor.onFetchTask = jest.fn().mockResolvedValue({ type: "continue" });
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "task/fetch",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({
                type: "invokeResponse",
                value: { body: { task: { type: "continue" } }, status: 200 }
            });

            expect(next).not.toBeCalled();
        });

        it("Should call onFetchTask, returning a message message", async () => {
            processor.onFetchTask = jest.fn().mockResolvedValue({ type: "message" });
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "task/fetch",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({
                type: "invokeResponse",
                value: { body: { task: { type: "message" } }, status: 200 }
            });

            expect(next).not.toBeCalled();
        });

        it("Should handle onFetchTask error", async () => {
            processor.onFetchTask = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "task/fetch",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should call onFetchTask - with invalid command id", async () => {
            processor.onFetchTask = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "task/fetch",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should not call onFetchTask - missing onQuery method", async () => {
            processor.onFetchTask = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "task/fetch",
                    value: {
                        commandId: "command"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(next).toBeCalled();
        });

        it("Should not call onFetchTask - not correct activity name", async () => {
            processor.onFetchTask = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onFetchTask).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });

    describe("onSubmitAction", () => {
        it("Should call onSubmitAction", async () => {
            processor.onSubmitAction = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onSubmitAction).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should handle onSubmitAction error", async () => {
            processor.onSubmitAction = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onSubmitAction).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should not call onSubmitAction - invalid command id", async () => {
            processor.onSubmitAction = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "wrong"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onSubmitAction).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });

        it("Should not call onSubmitAction - missing onSubmitAction method", async () => {
            processor.onSubmitAction = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "command"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(sendActivity).not.toBeCalled();
            expect(next).toBeCalled();
        });

        it("Should not call onSubmitAction - not correct activity name", async () => {
            processor.onSubmitAction = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onSubmitAction).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });

    describe("onBotMessagePreviewEdit ", () => {
        it("Should call onBotMessagePreviewEdit ", async () => {
            processor.onBotMessagePreviewEdit = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "command",
                        botMessagePreviewAction: "edit"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onBotMessagePreviewEdit).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should handle onBotMessagePreviewEdit  error", async () => {
            processor.onBotMessagePreviewEdit = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "command",
                        botMessagePreviewAction: "edit"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onBotMessagePreviewEdit).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should not call onBotMessagePreviewEdit  - invalid command id", async () => {
            processor.onBotMessagePreviewEdit = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "wrong",
                        botMessagePreviewAction: "edit"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onBotMessagePreviewEdit).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });

        it("Should not call onBotMessagePreviewEdit  - missing onBotMessagePreviewEdit  method", async () => {
            processor.onBotMessagePreviewEdit = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "command",
                        botMessagePreviewAction: "edit"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(sendActivity).not.toBeCalled();
            expect(next).toBeCalled();
        });

        it("Should not call onBotMessagePreviewEdit  - not correct activity name", async () => {
            processor.onBotMessagePreviewEdit = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onBotMessagePreviewEdit).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });

    describe("onBotMessagePreviewSend ", () => {
        it("Should call onBotMessagePreviewSend ", async () => {
            processor.onBotMessagePreviewSend = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "command",
                        botMessagePreviewAction: "send"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onBotMessagePreviewSend).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(next).not.toBeCalled();
        });

        it("Should handle onBotMessagePreviewSend  error", async () => {
            processor.onBotMessagePreviewSend = jest.fn().mockRejectedValue("error");
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "command",
                        botMessagePreviewAction: "send"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onBotMessagePreviewSend).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledTimes(1);
            expect(sendActivity).toBeCalledWith({ type: "invokeResponse", value: { body: "error", status: 500 } });
            expect(next).not.toBeCalled();
        });

        it("Should not call onBotMessagePreviewSend  - invalid command id", async () => {
            processor.onBotMessagePreviewSend = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "wrong",
                        botMessagePreviewAction: "send"
                    }
                }
            }), next);
            expect(result).toBe(undefined);
            expect(processor.onBotMessagePreviewSend).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });

        it("Should not call onBotMessagePreviewSend  - missing onBotMessagePreviewSend  method", async () => {
            processor.onBotMessagePreviewSend = undefined;
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({
                activity: {
                    name: "composeExtension/submitAction",
                    value: {
                        commandId: "command",
                        botMessagePreviewAction: "send"
                    }
                },
                sendActivity
            }), next);
            expect(result).toBe(undefined);
            expect(sendActivity).not.toBeCalled();
            expect(next).toBeCalled();
        });

        it("Should not call onBotMessagePreviewSend  - not correct activity name", async () => {
            processor.onBotMessagePreviewEdit = jest.fn().mockResolvedValue({});
            const mw = new exp.MessagingExtensionMiddleware("command", processor);

            const result = await mw.onTurn(stub<TurnContext>({}), next);
            expect(result).toBe(undefined);
            expect(processor.onBotMessagePreviewSend).toBeCalledTimes(0);
            expect(next).toBeCalled();
        });
    });
});
