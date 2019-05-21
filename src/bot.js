'use strict';

module.exports.setup = function(app) {
    var builder = require('botbuilder');
    var teams = require('botbuilder-teams');
    var config = require('config');

    if (!config.has("bot.appId")) {
        // We are running locally; fix up the location of the config directory and re-intialize config
        process.env.NODE_CONFIG_DIR = "../config";
        delete require.cache[require.resolve('config')];
        config = require('config');
    }
    // Create a connector to handle the conversations
    var connector = new teams.TeamsChatConnector({
        // It is a bad idea to store secrets in config files. We try to read the settings from
        // the config file (/config/default.json) OR then environment variables.
        // See node config module (https://www.npmjs.com/package/config) on how to create config files for your Node.js environment.
        appId: config.get("bot.appId"),
        appPassword: config.get("bot.appPassword")
    });
    
    var inMemoryBotStorage = new builder.MemoryBotStorage();
    
    // Define a simple bot with the above connector that echoes what it received
    var bot = new builder.UniversalBot(connector, function(session) {
        // Message might contain @mentions which we would like to strip off in the response
        // var text = teams.TeamsMessage.getTextWithoutMentions(session.message);
        // session.send('You said: %s', text);

        /*
        const msg = new teams.TeamsMessage(session).text('This is a test notification message.');
        // this is a dictionary which could be merged with other properties
        const alertFlag = teams.TeamsMessage.alertFlag;
        const notification = msg.sourceEvent({'*': alertFlag});

        // this should trigger an alert
        session.send(notification);
        */

       var toMention = {
        name: 'Rooparam Choudhary',
        id: 'a784a409-94c6-4341-b765-f36d949c07af'
        };
      var msg = new teams.TeamsMessage(session).text(teams.TeamsMessage.getTenantId(session.message));
      var mentionedMsg = msg.addMentionToText(toMention);
      var generalMessage = mentionedMsg.routeReplyToGeneralChannel();
      session.send(generalMessage);
    }).set('storage', inMemoryBotStorage);

    /*

    bot.dialog('/', [
        session => { builder.Prompts.choice(session, "Choose an option:", 'Route message to general channel|NotificationFeed'); },
        (session, results) => {
            switch (results.response.index) {
                case 0:
                    session.beginDialog('RouteMessageToGeneral');
                    break;
                case 1:
                    session.beginDialog('NotificationFeed');
                    break;
                default:
                    session.endDialog();
                    break;
            }
        }
    ]);

    bot.dialog('RouteMessageToGeneral', session => {
        // user name/user id
        var toMention = {
          name: 'Rooparam Choudhary',
          id: 'a784a409-94c6-4341-b765-f36d949c07af'
        };
        var msg = new teams.TeamsMessage(session).text(teams.TeamsMessage.getTenantId(session.message));
        var mentionedMsg = msg.addMentionToText(toMention);
        var generalMessage = mentionedMsg.routeReplyToGeneralChannel();
        session.send(generalMessage);
        session.endDialog();
      });

    bot.dialog('NotificationFeed', session => {
        const msg = new teams.TeamsMessage(session).text('This is a test notification message.');
        // this is a dictionary which could be merged with other properties
        const alertFlag = teams.TeamsMessage.alertFlag;
        const notification = msg.sourceEvent({'*': alertFlag});

        // this should trigger an alert
        session.send(notification);
        session.endDialog();
    });

    */

    // Setup an endpoint on the router for the bot to listen.
    // NOTE: This endpoint cannot be changed and must be api/messages
    app.post('/api/messages', connector.listen());

    // Export the connector for any downstream integration - e.g. registering a messaging extension
    module.exports.connector = connector;
};
