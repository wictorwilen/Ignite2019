import { BotDeclaration, MessageExtensionDeclaration, IBot, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes } from "botbuilder";
import PlannerTasksMessageExtension from "../plannerTasksMessageExtension/PlannerTasksMessageExtension";
import CreateTaskMessageExtension from "../createTaskMessageExtension/CreateTaskMessageExtension";
import { TeamsContext, TeamsActivityProcessor } from "botbuilder-teams";

// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for Planner Tasks Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)

export class PlannerTasksBot implements IBot {
    private readonly conversationState: ConversationState;
    /**
     * Local property for CreateTaskMessageExtension
     */
    @MessageExtensionDeclaration("createTaskMessageExtension")
    private _createTaskMessageExtension: CreateTaskMessageExtension;
    /**
     * Local property for PlannerTasksMessageExtension
     */
    @MessageExtensionDeclaration("plannerTasksMessageExtension")
    private _plannerTasksMessageExtension: PlannerTasksMessageExtension;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;
    private readonly activityProc = new TeamsActivityProcessor();

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        // Message extension CreateTaskMessageExtension
        this._createTaskMessageExtension = new CreateTaskMessageExtension();

        // Message extension PlannerTasksMessageExtension
        this._plannerTasksMessageExtension = new PlannerTasksMessageExtension();

        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);


    }

    /**
     * The Bot Framework `onTurn` handlder.
     * The Microsoft Teams middleware for Bot Framework uses a custom activity processor (`TeamsActivityProcessor`)
     * which is configured in the constructor of this sample
     */
    public async onTurn(context: TurnContext): Promise<any> {
        // transfer the activity to the TeamsActivityProcessor
        await this.activityProc.processIncomingActivity(context);
    }

}
