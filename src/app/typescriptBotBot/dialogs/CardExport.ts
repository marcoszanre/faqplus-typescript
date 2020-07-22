const WelcomeCard = require("./WelcomeCard.json");
const HelpCard = require("./HelpCard.json");

const ProactiveCard = {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "TextBlock",
                            "size": "medium",
                            "weight": "bolder",
                            "text": "${text}"
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                {
                                    "title": "De",
                                    "value": "${from}"
                                },
                                {
                                    "title": "Titulo",
                                    "value": "${titleitem}"
                                },
                                {
                                    "title": "Descrição",
                                    "value": "${descriptionitem}"
                                }
                            ]
                        }
                    ],
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.2",
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "Chat with ${from}",
                            "url": "https://teams.microsoft.com/l/chat/0/0?users=adelev@m365x165753.onmicrosoft.com&topicName=Help%20Bot&message=Re:"
                        },
                        {
                            "type": "Action.ShowCard",
                            "title": "Alterar Status",
                            "card": {
                                "type": "AdaptiveCard",
                                "version": "1.2",
                                "body": [
                                {
                                    "type": "Input.ChoiceSet",
                                    "id": "inputChoice",
                                    "choices": [
                                    {
                                        "$data": "${choices}",
                                        "title": "${title}",
                                        "value": "${value}"
                                    }
                            ],
                            "placeholder": "Alterar Status"
                        }
                                ],
                                "actions": [
                                    {
                    "type": "Action.Submit",
                    "title": "Alterar",
                    "data": {
                        "ticketID": "${ticketID}",
                        "ticketTitle": "${titleitem}",
                        "ticketDescription": "${descriptionitem}"
                    }
                    }
                                ]
                            }
                        }
                    ]
};

const templateNotifyCard =
        {          
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "Status de seu pedido de ajuda"
                },
                {
                    "type": "FactSet",
                    "facts": [
                        {
                            "title": "De",
                            "value": "${from}"
                        },
                        {
                            "title": "Titulo",
                            "value": "${title}"
                        },
                        {
                            "title": "Descrição",
                            "value": "${description}"
                        },
                        {
                            "title": "Status",
                            "value": "${status}"
                        }
                    ]
                }
            ],
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.2"
};

export {
    WelcomeCard,
    HelpCard,
    ProactiveCard,
    templateNotifyCard
}

