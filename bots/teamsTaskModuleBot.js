/* eslint-disable comma-dangle */
/* eslint-disable quotes */
/* eslint-disable indent */
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, MessageFactory, CardFactory } = require("botbuilder");
const { TaskModuleUIConstants } = require("../models/taskmoduleuiconstants");
const { TaskModuleIds } = require("../models/taskmoduleids");
const { TaskModuleResponseFactory } = require("../models/taskmoduleresponsefactory");

const Actions = [TaskModuleUIConstants.CustomForm];

class TeamsTaskModuleBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.baseUrl = process.env.BaseUrl;

    // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
    this.onMessage(async (context, next) => {
      // This displays two cards: A HeroCard. When any of the options are selected, `handleTeamsTaskModuleFetch`
      // is called.
      const reply = MessageFactory.list([this.getTaskModuleHeroCardOptions()]);
      await context.sendActivity(reply);

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });
  }

  handleTeamsTaskModuleFetch(context, taskModuleRequest) {
    // Called when the user selects an options from the displayed HeroCard or
    // AdaptiveCard.  The result is the action to perform.

    const cardTaskFetchValue = taskModuleRequest.data.data;
    var taskInfo = {}; // TaskModuleTaskInfo

    if (cardTaskFetchValue === TaskModuleIds.CustomForm) {
      // Display the CustomForm.html page, and post the form data back via
      // handleTeamsTaskModuleSubmit.
      taskInfo.url = taskInfo.fallbackUrl = this.baseUrl + "/" + TaskModuleIds.CustomForm + ".html";
      this.setTaskInfo(taskInfo, TaskModuleUIConstants.CustomForm);
    }

    return TaskModuleResponseFactory.toTaskModuleResponse(taskInfo);
  }

  async handleTeamsTaskModuleSubmit(context, taskModuleRequest) {
    // Called when data is being returned from the selected option (see `handleTeamsTaskModuleFetch').

    // Echo the users input back.  In a production bot, this is where you'd add behavior in
    // response to the input.
    await context.sendActivity(MessageFactory.text("Result: " + JSON.stringify(taskModuleRequest.data)));

    // Return TaskModuleResponse
    return {
      // TaskModuleMessageResponse
      task: {
        type: "message",
        value: "Thanks!",
      },
    };
  }

  setTaskInfo(taskInfo, uiSettings) {
    taskInfo.height = uiSettings.height;
    taskInfo.width = uiSettings.width;
    taskInfo.title = uiSettings.title;
  }

  getTaskModuleHeroCardOptions() {
    return CardFactory.heroCard(
      "Click on button below to select multiple dates:",
      "",
      null, // No images
      Actions.map((cardType) => {
        return {
          type: "invoke",
          title: cardType.buttonTitle,
          value: {
            type: "task/fetch",
            data: cardType.id,
          },
        };
      })
    );
  }
}

module.exports.TeamsTaskModuleBot = TeamsTaskModuleBot;
