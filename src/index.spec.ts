import { TurnContext } from "botbuilder-core";
import { stub } from "jest-auto-stub";

import * as exp from "./index";

const next = jest.fn().mockResolvedValue(undefined);

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

    it("Should not call onActionExecute", async () => {
        processor.onActionExecute = jest.fn().mockResolvedValue({});
        const mw = new exp.MessagingExtensionMiddleware("command", processor);

        const result = await mw.onTurn(stub<TurnContext>({}), next);
        expect(result).toBe(undefined);
        expect(processor.onActionExecute).toBeCalledTimes(0);
        expect(next).toBeCalled();
    });

    it("Should call onActionExecute", async () => {
        processor.onActionExecute = jest.fn().mockResolvedValue({});
        const mw = new exp.MessagingExtensionMiddleware("command", processor);

        const result = await mw.onTurn(stub<TurnContext>({ activity: { name: "adaptiveCard/action" } }), next);
        expect(result).toBe(undefined);
        expect(processor.onActionExecute).toBeCalledTimes(1);
        expect(next).not.toBeCalled();
    });

    it("Should not call onQuery", async () => {
        processor.onQuery = jest.fn().mockResolvedValue({});
        const mw = new exp.MessagingExtensionMiddleware("command", processor);

        const result = await mw.onTurn(stub<TurnContext>({}), next);
        expect(result).toBe(undefined);
        expect(processor.onQuery).toBeCalledTimes(0);
        expect(next).toBeCalled();
    });

});
