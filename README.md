# MSFTPP.Plugins.CustomAudit

## Custom Audit

## Overview

A Power Platform solution that will allow you to create rows in a Custom Audit table in Dataverse for enhanced analysis. 
This solution consists of a custom activity table, custom .NET plugin and a model driven app to configure and monitor the auditing. This is NOT a replacement for the OOTB auditing feature, but an enhancement to use in addition to that service.

_Disclaimer: This solution is authored by employees of Microsoft but is in no way endorsed by Microsoft Corporation. Code included is custom and can be modified by the receiving user and therefore is NOT supported by Microsoft or any of its affiliates or partners. No warranties or support are expressed or implied._

To Download the latest Solution, click the [Releases](https://github.com/InformedPowerPlatform/MSFTPP.Plugins.CustomAudit/releases) page.

## Prerequisites

### Requirements

- Microsoft Power Platform Environment with Dataverse data store enabled

In order to audit a table, the table MUST be enabled for **Creating a new activity**. This can be enabled by editing the properties of the table in the Power Apps maker tool.

>Note: Ensure you are in the correct environment at the top right of the screen.

- Select **Tables** in the left navigation.
- Locate the table you wish to audit and click on the name.
- Choose **Properties** to open the table properties screen.
- Click **Advanced options** to expand the options.
- Scroll to the **Make this table an option when** section and ensure that the **Creating a new activity** is checked.
- Click **Save** to update the properties.

## Contents of Solution

After importing the solution, a number of items are now available in your environment.

- A new Activity type table is created called **Custom Audit**. This table will store data in these columns.
    - **Table Name** : Friendly name of the table the audit row applies to
    - **Column Name** : Friendly name of the column the audit row applies to
    - **Changed On** : This is the standard _Created On_ column that is renamed to Changed On
    - **Changed By** : This is the standard _Owner_ column that is renamed to Changed By
    - **Action** : This is the action message sent by the plugin. Should be either **Create** or **Update**
    - **Regarding** : This is the standard _Regarding_ column for an activity row. It links back to the original record that was being audited.
    - **Old Value** : A text column that stores the old value of the audited row BEFORE the change.
    - **New Value** : A text column that stores the new value of the audited row BEFORE the change.
>Note: Old Value and New Value are limited to 4,000 characters. If you need to store more data, you will have to customize the table with a larger multi-line text column. You will also need to customize the .NET plugin to add the data to the new column.
- Model Driven App :: **Custom Audit**
    - This app is where you can configure the plugin using a custom HTML/JS Web Resource.
- Plugin-in assembly : **MSFTPP.Plugin.CustomAudit**
    - This is the custom .NET Plugin that handles the auditing of the configured tables and logs the data to the Custom Audit table.
- Web Resources
    - **msftpp\_CustomAuditBuilder.html** – HTML Front end UI for configuring the Audit Steps
    - **msftpp\_CustomAuditBuilder.js** – JS backend code that works with the HTML front end
    - SVG images for the Custom Audit table

## Configuring a table for audit

Once you have confirmed that the table you wish to audit is enabled for activities, you can now configure which columns will be audited.

- Launch the Power Apps maker tool and select the environment.
- From the Left Navigation, select Apps
- Locate the Model driven app called **Custom Audit Builder** and Click to Play this app.
- From the Left Navigation, select **Builder**.

### Adding a new Table/Message for audit

- To Add a New Table and Message (create/update) for auditing, click **+ Add New** button.
    - You will be presented with a list of tables. If the table you want to audit is not shown, recheck that the table has been enabled for activities. Only tables that are enabled for activities will show in the list.
- Select the Table to Audit
-  Select the Message of **Create** or **Update**.
    - The **Create** message will audit when a new row is added to the table and put all of the New Values for columns you request to be audited. This can be used as a baseline set of audit rows for reporting.
    - The **Update** message will audit when a row is updated with Old and New Values. If a column did not previously have a value (e.g. is was NULL), then the Old Value will show as blank and the New Value will contain what was added.
- Once you select a message, you will be presented with a list of all auditable columns in the table. Select the checkbox next to the columns you wish to audit.
> Note: Some standard columns are not presented (e.g. Created On, Modified On, etc) as they should not be audited.
    - As a best practice it is recommended that you DO NOT select ALL columns to be audited. Be selective about what really needs to be logged.
- After selecting the columns, click the **Submit** button at the top right side of the page. This will create the appropriate Steps and Images required in the plugin registration to audit the table.

### Edit and Existing Plugin Step

- If you have previously configured a plugin step to audit a table and wish to modify that step (e.g. add/remove columns), you can select the **Edit Existing** button **.**
- You will be presented with a list of the previously configured steps for the Audit plugin. Select the step you wish to modify.
- A list of columns is built showing ones that have already been selected for audit. Add or Remove columns by checking or unchecking the box next to the column name.
- After selecting the columns, click the **Submit** button at the top right side of the page. This will update the Plugin and make the changes.

## Manual Plugin Registration Steps to add a table for auditing

- Locate and launch the **XRMToolbox** and Connect to the environment you wish to update.
- Load the **Plugin Registration** plugin
- In the Tree, locate the **(Assembly) MSFTPP.Plugin.CustomAudit** and expand it.
- Expand the **(Plugin) MSFTPP.Plugin.CustomAudit** node.
- Right Click on the **(Plugin) MSFTPP.Plugin.CustomAudit** node and select **Register New Step**
    - **Message** : This will either be **Update** or **Create** , depending on which action you are auditing.
    - **Primary Entity** : This will be the logical/sdk name for the entity being audited.
        - e.g. account
    - **Secondary Entity** : Leave this blank
    - **Filtering Attributes** : This is where you select which columns being updated will fire the plugin. It is recommended that you click the ellipsis button, deselect all, and then select ONLY the columns that you wish to audit from the list.
    - Event Handler, Name, Run in Users Context, Execution Order should be defaults
    - **Eventing Pipeline Stage** : Post-operation
    - **Execution Mode** : Asynchronous
        - Also check the box next to **Delete AsyncOperation if StatusCode = Successful**
    - **Optional** : In the Unsecure Configuration box, you can put **traceon=true** to help trace issues in the plugin trace log. In production you can leave this out, or set to **traceon=false**
    - Click **Register new step** to save your changes.

- Right-Click on the new (Step) you just created and select **Register New Image**
    - **Image Type** :
        - If this is an Update message, select both Pre Image and Post Image
        - If this is a Create message, only Post Image should be available
    - **Name** : Type the word **Target** (case sensitive)
    - **Entity Alias** : Type the word **Target** (case sensitive)
    - **Parameters** : Click the Ellipsis to access a list of columns. This is where you select which columns will be passed in for auditing. It is recommended that you click the ellipsis button, deselect all, and then select ONLY the columns that you wish to audit from the list. If you want a column audited, it MUST be in this list. Ideally, this list should match the Filtering attributes you selected in a previous step.
    - Click **Register Step** to save the changes.

