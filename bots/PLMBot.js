// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Poc: This will be the single entry point for the Chatbot.
// Based on what is evaluated as the usecase, the control
// gets routed to the appropriate dialog in dialogs folder.

// Import required Bot Framework classes.
const { ActionTypes, ActivityHandler, CardFactory } = require('botbuilder');

const wipAdaptiveCard = require('../resources/wip.json');

// Welcomed User property name
const WELCOMED_USER = 'welcomedUserProperty';

class PLMBot extends ActivityHandler {
    /**
     *
     * @param {ConversationState} conversationState
     * @param {UserState} userState
     * @param {Dialog} dialog
     */
    constructor(conversationState, userState) {
        super();
        if (!conversationState) throw new Error('[DialogBot]: Missing parameter. conversationState is required');
        if (!userState) throw new Error('[DialogBot]: Missing parameter. userState is required');

        this.conversationState = conversationState;
        this.userState = userState;
        this.dialogState = this.conversationState.createProperty('DialogState');

        // Creates a new user property accessor.
        // See https://aka.ms/about-bot-state-accessors to learn more about the bot state and state accessors.
        this.welcomedUserProperty = userState.createProperty(WELCOMED_USER);

        this.userState = userState;

        this.onMessage(async (context, next) => {
			// PoC: onMessage is an event that gets triggered whenever a new user input is received.
            // Read UserState. If the 'DidBotWelcomedUser' does not exist (first time ever for a user)
            // set the default to false.
            

            // PoC: When the user clicks on Submit on an Adaptive Card, it would come up as a message here.
            // To distinguish between normal messages and the user entered information messages, we use the
            // below check.
            if (context.activity.text === undefined && context.activity.value ) {
                const scenario = context.activity.value.scenario;

                switch(scenario)
                {
                    case 'add user to group':
                        const { AddUserToGroupDialog } = require('../dialogs/addUserToGroupDialog.js');
                        const addUserToGroupDialog = new AddUserToGroupDialog(userState);
                        this.dialog = addUserToGroupDialog;

                        // PoC: Run the dialog that is responsible for Adding a user to the Group.
                        await this.dialog.run(context, this.dialogState );

                        // PoC: Let the flow proceed with the next step in the waterfall as needed.
                        await next();

                        break;

                    case 'change ownership':
                        const { ChangeOwnershipDialog } = require('../dialogs/changeOwnershipDialog.js');
                        const changeOwnershipDialog = new ChangeOwnershipDialog(userState);
                        this.dialog = changeOwnershipDialog;
    
                        // PoC: Run the dialog that is responsible for Change Ownership.
                        await this.dialog.run(context, this.dialogState );
    
                        // PoC: Let the flow proceed with the next step in the waterfall as needed.
                        await next();

                        break;
                }
                


            }
            else
            {
                // This example uses an exact match on user's input utterance.
                // Consider using LUIS or QnA for Natural Language Processing.
                const text = context.activity.text.toLowerCase();
                // PoC TODO: If a LUIS Configuration is indeed introduced, it would need to be evaluated on 'text' to understand the intent.
                switch (text) {
                case 'add user to group':

                    const AddUserToGroupAdaptiveCard = require('../resources/addUserToGroupDetailsAdaptiveCard.json');
                    await context.sendActivity({
                        attachments: [CardFactory.adaptiveCard(AddUserToGroupAdaptiveCard)]
                    });

                    break;
                case 'change ownership':

                    const changeOwnershipAdaptiveCard = require('../resources/changeOwnershipDetailsAdaptiveCard.json');
                    await context.sendActivity({
                        attachments: [CardFactory.adaptiveCard(changeOwnershipAdaptiveCard)]
                    });

                    break;
                default:
                    await context.sendActivity({
                        attachments: [CardFactory.adaptiveCard(wipAdaptiveCard)]
                    });
                }
            }

            // PoC: next() is needed to keep the waterfall alive.
            await next();
        });

        // Sends welcome messages to conversation members when they join the conversation.
        // Messages are only sent to conversation members who aren't the bot.
        this.onMembersAdded(async (context, next) => {

            // PoC: onMembersAdded is an event that gets triggered when an end user or the bot joins the conversation.
            for (const idx in context.activity.membersAdded) {
                // PoC: The condition below is needed to only send the welcome message when the user joins the conversation.
                if (context.activity.membersAdded[idx].id !== context.activity.recipient.id) {
                    await context.sendActivity(`Hello There!. I am the PLM Chatbot and I can assist you with your problem.`);
                }
            }

            // PoC: next() is needed to keep the waterfall alive.
            await next();
        });
    }

    /**
     * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
     */
    async run(context) {
        await super.run(context);

        // Save state changes
        await this.userState.saveChanges(context);
    }
}

module.exports.PLMBot = PLMBot;
