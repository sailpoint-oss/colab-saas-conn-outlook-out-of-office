import {
    Context,
    createConnector,
    readConfig,
    Response,
    logger,
    StdAccountListOutput,
    StdAccountListInput,
    StdAccountReadInput,
    StdAccountReadOutput,
    StdTestConnectionOutput,
} from '@sailpoint/connector-sdk'
import { MyClient } from './my-client'

// Connector must be exported as module property named connector
export const connector = async () => {
    // Get connector source config
    const config = await readConfig()

    // Use the vendor SDK, or implement own client as necessary, to initialize a client
    const myClient = new MyClient(config)

    return createConnector()
        .stdTestConnection(async (context: Context, input: undefined, res: Response<StdTestConnectionOutput>) => {
            logger.info('Running test connection')
            await myClient.testConnection()
            res.send({})
        })
        .stdAccountList(async (context: Context, input: StdAccountListInput, res: Response<StdAccountListOutput>) => {
            const accounts = await myClient.getAllAccounts()

            for (const account of accounts) {
                //console.log('This is my account ' + account.get('id'))
                res.send({
                    identity: <string>account.get('id'),
                    uuid: <string>account.get('userPrincipalName'),
                    attributes: {
                        id: <string>account.get('id'),
                        mail: <string>account.get('mail'),
                        displayName: <string>account.get('displayName'),
                        userPrincipalName: <string>account.get('userPrincipalName'),
                        firstname: <string>account.get('firstName'),
                        lastname: <string>account.get('lastName'),
                        automaticRepliesSetting: <string>account.get('automaticRepliesSetting'),
                        lastPasswordChangeDateTime: <string>account.get('lastPasswordChangeDateTime'),
                        scheduledEndDateTime: <string>account.get('scheduledEndDateTime'),
                        scheduledStartDateTime: <string>account.get('scheduledStartDateTime'),
                        automaticRepliesStatus: <string>account.get('automaticRepliesStatus'),
                    },
                })
            }
            logger.info(`stdAccountList sent ${accounts.length} accounts`)
        })
        .stdAccountRead(async (context: Context, input: StdAccountReadInput, res: Response<StdAccountReadOutput>) => {
            const account = await myClient.getAccount(input.identity)

            res.send({
                identity: <string>account.get('id'),
                uuid: <string>account.get('userPrincipalName'),
                attributes: {
                    id: <string>account.get('id'),
                    mail: <string>account.get('mail'),
                    displayName: <string>account.get('displayName'),
                    userPrincipalName: <string>account.get('userPrincipalName'),
                    firstname: <string>account.get('firstName'),
                    lastname: <string>account.get('lastName'),
                    automaticRepliesSetting: <string>account.get('automaticRepliesSetting'),
                    lastPasswordChangeDateTime: <string>account.get('lastPasswordChangeDateTime'),
                    scheduledEndDateTime: <string>account.get('scheduledEndDateTime'),
                    scheduledStartDateTime: <string>account.get('scheduledStartDateTime'),
                    automaticRepliesStatus: <string>account.get('automaticRepliesStatus'),
                },
            })
            logger.info(`stdAccountRead read account : ${input.identity}`)
        })
}
