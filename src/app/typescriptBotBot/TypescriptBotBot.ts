import { BotDeclaration, MessageExtensionDeclaration, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, TeamsActivityHandler, MessageFactory, ConversationReference } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import { WelcomeCard, HelpCard, ProactiveCard, templateNotifyCard } from "./dialogs/CardExport";
import * as ACData from "adaptivecards-templating";
import { initTableSvc, insertTicket, updatetTicket, getTicket, ITicket } from './tableService';
// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for typescript bot Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)

export class TypescriptBotBot extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();
        
        // Init Table Service
        initTableSvc();
        
        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));

        // Set up the Activity processing

        this.onTurn(async (context: TurnContext, next): Promise<void> => {
            const isValidTenant = await this.checkTenant(context.activity.conversation.tenantId);
            if (isValidTenant) {
                // tenant autorizado
                await context.sendActivity('tenant autorizado!');
                await next();
            } else { 
                // invalid tenant
                await context.sendActivity('invalid tenant');
                return;
            }
        });


        this.onMessage(async (context: TurnContext): Promise<void> => {
            

            switch (context.activity.conversation.conversationType) {
                case 'personal':
                    if (context.activity.value) {
                        // To-Do - add blank logic

                        // Send Typing Activity
                        await context.sendActivity({type:  ActivityTypes.Typing});

                        // adaptive card submission
                        const teamsChannelId = process.env.TEAMS_CHANNEL_ID;
                        const ticketID = await this.generateGUID();
                        const adaptiveNotifyCard = await this.buildProactiveNotify(context, ticketID, "abrir");
                        const notifyCard = CardFactory.adaptiveCard(adaptiveNotifyCard);
                        const message = MessageFactory.attachment(notifyCard);
                        await this.teamsCreateConversation(context, teamsChannelId, message);
                        
                        // Send user proactive message and store conversationreference
                        const conversationReference = TurnContext.getConversationReference(context.activity);
                        await context.adapter.continueConversation(conversationReference, async turnContext => {
                            await this.sendCardNotify(context, "Created");
                        });

                        // Insert Ticket to table storage
                        await insertTicket(
                            context,
                            ticketID,
                            JSON.stringify(conversationReference)
                        );

                        await context.sendActivity("Vou reportar agora no canal, obrigado!");
                    }
                    else {
                        // default message submission
                        let text = TurnContext.removeRecipientMention(context.activity);
                        text = text.toLowerCase();
                        text = text.trim();

                        // Send Typing Activity
                        await context.sendActivity({type:  ActivityTypes.Typing});
                        
                        if (text == "help") {
                            // Send HelpCard
                            const helpCard = CardFactory.adaptiveCard(HelpCard);
                            await context.sendActivity({ attachments: [helpCard] });
                        } else {
                            // Process QnA Logic
                            await context.sendActivity('Mensagem de teste');
                        }
                    }
                    break;
                case 'channel':
                    if (context.activity.value) {
                        const replyTxt = `${context.activity.from.name} alterou o status do ${context.activity.value.ticketID} para '${context.activity.value.inputChoice}'`;
                        await context.sendActivity(replyTxt);

                        switch (context.activity.value.inputChoice) {
                            case 'atribuir a mim':
                                // Send Typing Activity
                                await context.sendActivity({type:  ActivityTypes.Typing});

                                // Atualizar ticket para atribuido
                                await updatetTicket(context.activity.value.ticketID, 'atribuido');
                                
                                // Notify User
                                await getTicket(context.activity.value.ticketID).then(async (ticket: ITicket) => {
                                    const notifyConversationReference: Partial<ConversationReference> = JSON.parse(ticket.conversationid);
                                    await context.adapter.continueConversation(notifyConversationReference, async turnContext => {
                                        const TemplateNotifyCard = await this.buildCardNotify(ticket.createdBy, ticket.title, ticket.description, "Atribuido");
                                        await turnContext.sendActivity({ attachments: [TemplateNotifyCard] });
                                    });
                                });

                                // Atualizar cartão
                                const adaptiveNotifyCard = await this.buildProactiveNotify(context, context.activity.value.ticketID, "atribuir");
                                const notifyCard = CardFactory.adaptiveCard(adaptiveNotifyCard);
                                const message = MessageFactory.attachment(notifyCard);
                                message.id = context.activity.replyToId;
                                await context.updateActivity(message);
                                break;

                            case 'fechar':
                                // Send Typing Activity
                                await context.sendActivity({type:  ActivityTypes.Typing});

                                // Atualizar ticket para atribuido
                                await updatetTicket(context.activity.value.ticketID, 'fechado');

                                // Notify User
                                await getTicket(context.activity.value.ticketID).then(async (ticket: ITicket) => {
                                    const notifyConversationReference: Partial<ConversationReference> = JSON.parse(ticket.conversationid);
                                    await context.adapter.continueConversation(notifyConversationReference, async turnContext => {
                                        const TemplateNotifyCard = await this.buildCardNotify(ticket.createdBy, ticket.title, ticket.description, "Fechado");
                                        await turnContext.sendActivity({ attachments: [TemplateNotifyCard] });
                                    });
                                });

                                // Atualizar cartão
                                const fecharAdaptiveNotifyCard = await this.buildProactiveNotify(context, context.activity.value.ticketID, "fechar");
                                const fecharNotifyCard = CardFactory.adaptiveCard(fecharAdaptiveNotifyCard);
                                const fecharMessage = MessageFactory.attachment(fecharNotifyCard);
                                fecharMessage.id = context.activity.replyToId;
                                await context.updateActivity(fecharMessage);
                                break;

                            case 'reabrir':
                                // Send Typing Activity
                                await context.sendActivity({type:  ActivityTypes.Typing});

                                // Atualizar ticket para atribuido
                                await updatetTicket(context.activity.value.ticketID, 'aberto');

                                // Notify User
                                await getTicket(context.activity.value.ticketID).then(async (ticket: ITicket) => {
                                    const notifyConversationReference: Partial<ConversationReference> = JSON.parse(ticket.conversationid);
                                    await context.adapter.continueConversation(notifyConversationReference, async turnContext => {
                                        const TemplateNotifyCard = await this.buildCardNotify(ticket.createdBy, ticket.title, ticket.description, "Aberto");
                                        await turnContext.sendActivity({ attachments: [TemplateNotifyCard] });
                                    });
                                });

                                // Atualizar cartão
                                const reabrirAdaptiveNotifyCard = await this.buildProactiveNotify(context, context.activity.value.ticketID, "reabrir");
                                const reabrirNotifyCard = CardFactory.adaptiveCard(reabrirAdaptiveNotifyCard);
                                const reabrirMessage = MessageFactory.attachment(reabrirNotifyCard);
                                reabrirMessage.id = context.activity.replyToId;
                                await context.updateActivity(reabrirMessage);
                                break;
                        }

                    }

                    break;
                case 'groupChat':

                    break;
                    default: 

                    break;
            }

            // Save state changes
            return this.conversationState.saveChanges(context);
        });

        

        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                        await context.sendActivity({ attachments: [welcomeCard] });
                    }
                }
            }
        });

        this.onMessageReaction(async (context: TurnContext): Promise<void> => {
            const added = context.activity.reactionsAdded;
            if (added && added[0]) {
                await context.sendActivity({
                    textFormat: "xml",
                    text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                });
            }
        });;
   }

   async teamsCreateConversation(context, teamsChannelId, message) {
    const conversationParameters = {
        isGroup: true,
        channelData: {
            channel: {
                id: teamsChannelId
            }
        },

        activity: message
    };

    const connectorClient = context.adapter.createConnectorClient(context.activity.serviceUrl);
    await connectorClient.conversations.createConversation(conversationParameters);
    }

   async teamsUserConversation(context, message) {
    const conversationParameters = {
        user: { id: context.activity.from.id },
        channelData: {
            tenant: {
                id: context.activity.conversation.tenantId
            }
        },

        activity: message
    };

    const connectorClient = context.adapter.createConnectorClient(context.activity.serviceUrl);
    await connectorClient.conversations.createConversation(conversationParameters);
    }

    async checkTenant(tenant) {
        return tenant === process.env.TENANT_ID;
    }

    async sendCardNotify(context, status: string ) {
        // Send First Notify Card
        let template = new ACData.Template(templateNotifyCard);
        let cardPayload = template.expand({
        $root: {
            from: context.activity.from.name,
            title: context.activity.value.txtTitle,
            description: context.activity.value.txtDescription,
            status: status
        }
        });
        const TemplateNotifyCard = CardFactory.adaptiveCard(cardPayload);
        await context.sendActivity({ attachments: [TemplateNotifyCard] });
    }

    async buildCardNotify(from: string, title: string, description: string, status: string ) {
        // Send First Notify Card
        let template = new ACData.Template(templateNotifyCard);
        let cardPayload = template.expand({
        $root: {
            from: from,
            title: title,
            description: description,
            status: status
        }
        });
        const TemplateNotifyCard = CardFactory.adaptiveCard(cardPayload);
        return TemplateNotifyCard;
    }

    async buildProactiveNotify(context, ticketID: string, type: string) {
        let choicesAvailable;
        let textValue;
        let titleText;
        let descriptionText;
        switch (type) {
            case "abrir":
                titleText = context.activity.value.txtTitle;
                descriptionText = context.activity.value.txtDescription;
                textValue = "Novo pedido de ajuda";
                choicesAvailable =
                [
                    {
                        "title": "atribuir a mim",
                        "value": "atribuir a mim"
                    },
                    {
                        "title": "fechar",
                        "value": "fechar"
                    }
                ]
                break;
            case "reabrir":
                titleText = context.activity.value.ticketTitle;
                descriptionText = context.activity.value.ticketDescription;
                textValue = "Novo pedido de ajuda";
                choicesAvailable =
                [
                    {
                        "title": "atribuir a mim",
                        "value": "atribuir a mim"
                    },
                    {
                        "title": "fechar",
                        "value": "fechar"
                    }
                ]
                break;
            case "atribuir":
                titleText = context.activity.value.ticketTitle;
                descriptionText = context.activity.value.ticketDescription;
                textValue = `Pedido atribuído à ${context.activity.from.name}`;
                choicesAvailable = 
                [
                    {
                        "title": "reabrir",
                        "value": "reabrir"
                    },
                    {
                        "title": "fechar",
                        "value": "fechar"
                    }
                ]
                break;
                case "fechar":
                titleText = context.activity.value.ticketTitle;
                descriptionText = context.activity.value.ticketDescription;
                textValue = `Pedido fechado por ${context.activity.from.name}`;
                choicesAvailable = 
                [
                    {
                        "title": "reabrir",
                        "value": "reabrir"
                    }
                ]
                break;
        }
        // Send First Notify Card
        let template = new ACData.Template(ProactiveCard);
        let cardPayload = template.expand({
        $root: {
            text: textValue,
            from: context.activity.from.name,
            titleitem: titleText,
            descriptionitem: descriptionText,
            ticketID: ticketID,
            choices: choicesAvailable
        }
        });
        return cardPayload;
    }

    generateGUID() {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
            var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }


}
