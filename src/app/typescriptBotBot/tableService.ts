import { TurnContext } from "botbuilder";

const azure = require('azure-storage');

const tableSvc = azure.createTableService(process.env.STORAGE_ACCOUNT_NAME, process.env.STORAGE_ACCOUNT_ACCESSKEY);

const initTableSvc = () => {
    tableSvc.createTableIfNotExists('ticketsTable', function(error){
        if(!error){
          // Table exists or created
          console.log('table service done');
        }
    });
}

const insertTicket = (context: TurnContext, ticketID: string, conversationid: string) => {
        let ticketInfo = {
            PartitionKey: {'_':'tickets'},
            RowKey: {'_': ticketID},
            title: {'_': context.activity.value.txtTitle},
            description: {'_': context.activity.value.txtDescription},
            createdBy: {'_': context.activity.from.name},
            conversationid: {'_': conversationid},
            status: {'_': 'aberto'}
        };

        tableSvc.insertEntity('ticketsTable',ticketInfo, function (error) {
            if(!error){
              // Entity inserted
              console.log('success!!!');
            } else {
                console.log(error);
            }
        });
}

const updatetTicket = (ticketID: string, status: string) => {
    let ticketInfo = {
        PartitionKey: {'_':'tickets'},
        RowKey: {'_': ticketID},
        status: {'_': status}
    };

    tableSvc.mergeEntity('ticketsTable',ticketInfo, function (error) {
        if(!error){
          // Entity inserted
          console.log('success!!!');
        } else {
            console.log(error);
        }
    });
}

const getTicket = async (ticketID: string) => {

    return new Promise((resolve) => {

        tableSvc.retrieveEntity('ticketsTable', 'tickets', ticketID, function (error, result) {
            if (!error) {
                // result contains the entity
                const ticketInfo: ITicket = {
                    title: result.title._,
                    description: result.description._,
                    conversationid: result.conversationid._,
                    status: result.status._,
                    createdBy: result.createdBy._
                };
                // console.log(returnData);
                resolve(ticketInfo);
            } 
        });
    });
}

interface ITicket {
    title: string;
    description: string;
    conversationid: string;
    status: string;
    createdBy: string;
}

export {
    initTableSvc,
    insertTicket,
    updatetTicket,
    getTicket,
    ITicket
}