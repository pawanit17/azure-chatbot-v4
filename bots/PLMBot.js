// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Import required Bot Framework classes.
const { ActionTypes, ActivityHandler, CardFactory } = require('botbuilder');

const AddUserToGroupAdaptiveCard = require('../resources/addUserToGroup.json');
const changeOwnershipAdaptiveCard = require('../resources/changeOwnership.json');
const wipAdaptiveCard = require('../resources/wip.json');

// Welcomed User property name
const WELCOMED_USER = 'welcomedUserProperty';

class PLMBot extends ActivityHandler {
    /**
     *
     * @param {UserState} User state to persist boolean flag to indicate
     *                    if the bot had already welcomed the user
     */
    constructor(userState) {
        super();
        // Creates a new user property accessor.
        // See https://aka.ms/about-bot-state-accessors to learn more about the bot state and state accessors.
        this.welcomedUserProperty = userState.createProperty(WELCOMED_USER);

        this.userState = userState;

        this.onMessage(async (context, next) => {
			// PoC: onMessage is an event that gets triggered whenever a new user input is received.
            // Read UserState. If the 'DidBotWelcomedUser' does not exist (first time ever for a user)
            // set the default to false.
            
			// This example uses an exact match on user's input utterance.
			// Consider using LUIS or QnA for Natural Language Processing.
			const text = context.activity.text.toLowerCase();
			switch (text) {
			case 'add user to group':
				await context.sendActivity({
                    attachments: [CardFactory.adaptiveCard(AddUserToGroupAdaptiveCard)]
                });
				break;
			case 'change ownership':
                await context.sendActivity({
                    attachments: [CardFactory.adaptiveCard(changeOwnershipAdaptiveCard)]
                });
				break;
			default:
				await context.sendActivity({
                    attachments: [CardFactory.adaptiveCard(wipAdaptiveCard)]
                });
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
