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
    var bot = new builder.UniversalBot(connector).set('storage', inMemoryBotStorage);

    // Strip bot at mention text, set text property to text without specific Bot at mention, find original text in textWithBotMentions
    // e.g. original text "<at>zel-bot-1</at> hello please find <at>Bot</at>" and zel-bot-1 is the Bot we at mentions. 
    // Then it text would be "hello please find <at>Bot</at>", the original text could be found at textWithBotMentions property.
    // This is to resolve inaccuracy for regex or LUIS scenarios.
    var stripBotAtMentions = new teams.StripBotAtMentions();
    bot.use(stripBotAtMentions);
    bot.dialog('/', [
        session => {
            builder.Prompts.choice(session, "Choose an option:", 'Fetch channel list|Mention user|Start new 1 on 1 chat|Route message to general channel|FetchMemberList|Send O365 actionable connector card|FetchTeamInfo(at Bot in team)|Start New Reply Chain (in channel)|Issue a Signin card to sign in a Facebook app|Logout Facebook app and clear cached credentials|MentionChannel|MentionTeam|NotificationFeed|Bot Delete Message');
        },
        (session, results) => {
            switch (results.response.index) {
                case 0:
                    session.beginDialog('FetchChannelList');
                    break;
                case 1:
                    session.beginDialog('MentionUser');
                    break;
                case 2:
                    session.beginDialog('StartNew1on1Chat');
                    break;
                case 3:
                    session.beginDialog('RouteMessageToGeneral');
                    break;
                case 4:
                    session.beginDialog('FetchMemberList');
                    break;
                case 5:
                    session.beginDialog('SendO365Card');
                    break;
                case 6:
                    session.beginDialog('FetchTeamInfo');
                    break;
                case 7:
                    session.beginDialog('StartNewReplyChain');
                    break;
                case 8:
                    session.beginDialog('Signin');
                    break;
                case 9:
                    session.beginDialog('Signout');
                    break;
                case 10:
                    session.beginDialog('MentionChannel');
                    break;
                case 11:
                    session.beginDialog('MentionTeam');
                    break;
                case 12:
                    session.beginDialog('NotificationFeed');
                    break;
                case 13:
                    session.beginDialog('BotDeleteMessage');
                    break;
                default:
                    session.endDialog();
                    break;
            }
        }
    ]);

    bot.on('conversationUpdate', function (message) {
        console.log(message);
        var event = teams.TeamsMessage.getConversationUpdateData(message);
    });
    bot.dialog('FetchChannelList', function (session) {
        var teamId = session.message.sourceEvent.team.id;
        connector.fetchChannelList(session.message.address.serviceUrl, teamId, function (err, result) {
            if (err) {
                session.endDialog('There is some error');
            }
            else {
                session.endDialog('%s', JSON.stringify(result));
            }
        });
    });
    bot.dialog('FetchMemberList', function (session) {
        var conversationId = session.message.address.conversation.id;
        connector.fetchMembers(session.message.address.serviceUrl, conversationId, function (err, result) {
            if (err) {
                session.endDialog('There is some error');
            }
            else {
                session.endDialog('%s', JSON.stringify(result));
            }
        });
    });
    bot.dialog('FetchTeamInfo', function (session) {
        var teamId = session.message.sourceEvent.team.id;
        connector.fetchTeamInfo(session.message.address.serviceUrl, teamId, function (err, result) {
            if (err) {
                session.endDialog('There is some error');
            }
            else {
                session.endDialog('%s', JSON.stringify(result));
            }
        });
    });
    bot.dialog('StartNewReplyChain', function (session) {
        var channelId = session.message.sourceEvent.channel.id;
        var message = new teams.TeamsMessage(session).text(teams.TeamsMessage.getTenantId(session.message));
        connector.startReplyChain(session.message.address.serviceUrl, channelId, message, function (err, address) {
            if (err) {
                console.log(err);
                session.endDialog('There is some error');
            }
            else {
                console.log(address);
                var msg = new teams.TeamsMessage(session).text("this is a reply message.").address(address);
                session.send(msg);
                session.endDialog();
            }
        });
    });
    bot.dialog('MentionUser', function (session) {
        // user name/user id
        var user = {
            id: userId,
            name: 'Bill Zeng'
        };
        var mention = new teams.UserMention(user);
        var msg = new teams.TeamsMessage(session).addEntity(mention).text(mention.text + ' ' + teams.TeamsMessage.getTenantId(session.message));
        session.send(msg);
        session.endDialog();
    });
    bot.dialog('MentionChannel', function (session) {
        // user name/user id
        var channelId = null;
        if (session.message.address.conversation.id) {
            var splitted = session.message.address.conversation.id.split(';', 1);
            channelId = splitted[0];
        }
        var teamId = session.message.sourceEvent.team.id;
        connector.fetchChannelList(session.message.address.serviceUrl, teamId, function (err, result) {
            if (err) {
                session.endDialog('There is some error');
            }
            else {
                var channelName = null;
                for (var i in result) {
                    var channelInfo = result[i];
                    if (channelId == channelInfo['id']) {
                        channelName = channelInfo['name'] || 'General';
                        break;
                    }
                }
                var channel = {
                    id: channelId,
                    name: channelName
                };
                var mention = new teams.ChannelMention(channel);
                var msg = new teams.TeamsMessage(session).addEntity(mention).text(mention.text + ' This is a test message to at mention the channel.');
                session.send(msg);
                session.endDialog();
            }
        });
    });
    bot.dialog('MentionTeam', function (session) {
        // user name/user id
        var channelId = null;
        if (session.message.address.conversation.id) {
            var splitted = session.message.address.conversation.id.split(';', 1);
            channelId = splitted[0];
        }
        var team = {
            id: channelId,
            name: 'All'
        };
        var mention = new teams.TeamMention(team);
        var msg = new teams.TeamsMessage(session).addEntity(mention).text(mention.text + ' This is a test message to at mention the team. ');
        session.send(msg);
        session.endDialog();
    });
    bot.dialog('NotificationFeed', function (session) {
        // user name/user id
        var msg = new teams.TeamsMessage(session).text("This is a test notification message.");
        // This is a dictionary which could be merged with other properties
        var alertFlag = teams.TeamsMessage.AlertFlag();
        var notification = msg.sourceEvent({
            'msteams': alertFlag
        });
        // this should trigger an alert
        session.send(notification);
        session.endDialog();
    });
    bot.dialog('StartNew1on1Chat', function (session) {
        var address = {
            channelId: 'msteams',
            user: { id: userId },
            channelData: {
                tenant: {
                    id: tenantId
                }
            },
            bot: {
                id: appId,
                name: appName
            },
            serviceUrl: session.message.address.serviceUrl,
            useAuth: true
        };
        bot.beginDialog(address, '/');
    });
    bot.dialog('BotDeleteMessage', function (session) {
        var msg = new teams.TeamsMessage(session).text("Bot will delete this message in 5 sec.");
        bot.send(msg, function (err, response) {
            if (err) {
                console.log(err);
                session.endDialog();
            }
            console.log('Proactive message response:');
            console.log(response);
            console.log('---------------------------------------------------');
            setTimeout(function () {
                var activityId = null;
                var messageAddress = null;
                if (response[0]) {
                    messageAddress = response[0];
                    activityId = messageAddress.id;
                }
                if (activityId == null) {
                    console.log('Message failed to send.');
                    session.endDialog();
                    return;
                }
                // Bot delete message
                var address = {
                    channelId: 'msteams',
                    user: messageAddress.user,
                    bot: messageAddress.bot,
                    id: activityId,
                    serviceUrl: session.message.address.serviceUrl,
                    conversation: {
                        id: session.message.address.conversation.id
                    }
                };
                connector["delete"](address, function (err) {
                    if (err) {
                        console.log(err);
                    }
                    else {
                        console.log("Message: " + activityId + " deleted successfully.");
                    }
                    // Try editing deleted message would fail
                    var newMsg = new builder.Message().address(address).text("To edit message.");
                    connector.update(newMsg.toMessage(), function (err, address) {
                        if (err) {
                            console.log(err);
                            console.log('Deleted message can not be edited.');
                        }
                        else {
                            console.log("There is something wrong. Message: " + activityId + " edited successfully.");
                            console.log(address);
                        }
                        session.endDialog();
                    });
                });
            }, 5000);
        });
    });
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

    // Setup an endpoint on the router for the bot to listen.
    // NOTE: This endpoint cannot be changed and must be api/messages
    app.post('/api/messages', connector.listen());

    // Export the connector for any downstream integration - e.g. registering a messaging extension
    module.exports.connector = connector;
};
