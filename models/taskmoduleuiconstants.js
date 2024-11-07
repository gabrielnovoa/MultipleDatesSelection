// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { UISettings } = require("./uisettings");
const { TaskModuleIds } = require("./taskmoduleids");

const TaskModuleUIConstants = {
  CustomForm: new UISettings(510, 450, "Select a Date", TaskModuleIds.CustomForm, "Select Dates"),
};

module.exports.TaskModuleUIConstants = TaskModuleUIConstants;
