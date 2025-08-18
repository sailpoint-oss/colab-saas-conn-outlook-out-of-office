import { ConnectorError } from '@sailpoint/connector-sdk'
import { Client } from '@microsoft/microsoft-graph-client'
import { ClientSecretCredential } from '@azure/identity'
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials'
import { PageCollection } from '@microsoft/microsoft-graph-client'
import { PageIterator } from '@microsoft/microsoft-graph-client'
import { PageIteratorCallback } from '@microsoft/microsoft-graph-client'
import 'isomorphic-fetch'


function mappingMailboxSettings(user: Map<string, string>, mailboxSettings: Map<string, string>) {
    user.set('automaticRepliesSetting', mailboxSettings.get('automaticRepliesSetting') as string)
    user.set('scheduledEndDateTime', mailboxSettings.get('scheduledEndDateTime') as string)
    user.set('scheduledStartDateTime', mailboxSettings.get('scheduledStartDateTime') as string)
    user.set('automaticRepliesStatus', mailboxSettings.get('automaticRepliesStatus') as string)
    return user;
}

function mappingUser(user: any) {
    const result = new Map()
    result.set('id', user.id)
    result.set('uuid', user.id)
    result.set('email', user.mail)
    result.set('displayName', user.displayName)
    result.set('userPrincipalName', user.userPrincipalName)
    result.set('firstName', user.firstName)
    result.set('lastName', user.lastName)
    result.set('lastPasswordChangeDateTime', user.lastPasswordChangeDateTime)
    return result
}

export class MyClient {
    private readonly client: Client
    private readonly filter: string
    private readonly pageSize: number

    constructor(config: any) {
        // Fetch necessary properties from config.
        // Following properties actually do not exist in the config -- it just serves as an example.

        if (config?.tenantId == null || config?.clientId == null || config?.clientSecret == null) {
            throw new ConnectorError('tenantId, clientId or clientSecret must be provided from config')
        }

        const credential = new ClientSecretCredential(config?.tenantId, config?.clientId, config?.clientSecret)
        const authProvider = new TokenCredentialAuthenticationProvider(credential, {
            scopes: ['https://graph.microsoft.com/.default'],
        })

        this.filter = config?.filter ?? ""
        this.pageSize = Math.min(config?.pageSize ?? 250, 999) // You can set the page size up to 999
        console.log("pageSize=" + this.pageSize);



        this.client = Client.initWithMiddleware({
            debugLogging: true,
            authProvider,
            // Use the authProvider object to create the class.
        })
    }

    async getMailboxSettings(identity: string): Promise<Map<string, string>> {
        let result: Map<string, string> = new Map()
        try {
            await this.client
                .api(`/users/${identity}/mailboxSettings/automaticRepliesSetting`)
                .top(1)
                //.select(['status', 'scheduledStartDateTime', 'scheduledEndDateTime'])
                .get()
                .then((mailboxSettings) => {
                    result.set('automaticRepliesSetting', mailboxSettings.status)
                    result.set('scheduledEndDateTime', mailboxSettings.scheduledEndDateTime.dateTime)
                    result.set('scheduledStartDateTime', mailboxSettings.scheduledStartDateTime.dateTime)

                    if (mailboxSettings.status == 'alwaysEnabled') {
                        result.set('automaticRepliesStatus', 'enabled')
                    }

                    if (mailboxSettings.status == 'scheduled') {
                        const start = new Date(mailboxSettings.scheduledStartDateTime.dateTime)
                        const end = new Date(mailboxSettings.scheduledEndDateTime.dateTime)
                        const now = new Date()
                        if (start < now && end > now) {
                            result.set('automaticRepliesStatus', 'enabled')
                        } else {
                            result.set('automaticRepliesStatus', 'disabled')
                        }
                    }

                    if (mailboxSettings.status == 'disabled') {
                        result.set('automaticRepliesStatus', 'disabled')
                    }
                })
        } catch (error) {
            console.error('Error fetching mailbox settings:', error)
        }
        return result
    }

    async *getAllAccounts(): AsyncGenerator<Map<string, string>> {
        let results: Map<string, string>[] = []
        let count = 0;
        let response: PageCollection = await this.client
            .api('/users')
            .header('ConsistencyLevel', 'eventual')
            .top(this.pageSize)
            .select([
                'id',
                'userPrincipalName',
                'mail',
                'displayName',
                'firstName',
                'lastName',
                'lastPasswordChangeDateTime',
            ])
            .filter(this.filter)
            .get()

        let callback: PageIteratorCallback = (user) => {
            const result = mappingUser(user)
            results.push(result)
            count++;
            // If we've iterated over the limit,
            // stop the iteration by returning false
            return count < this.pageSize;
        }

        // Creating a new page iterator instance with client a graph client
        // instance, page collection response from request and callback
        let pageIterator = new PageIterator(this.client, response, callback)

        // This iterates the collection until the nextLink is drained out.
        await pageIterator.iterate()
        // Refactoring to return more regularly user info
        // cd. https://learn.microsoft.com/en-us/graph/sdks/paging?tabs=typescript#stopping-and-resuming-the-iteration
        while (!pageIterator.isComplete()) {
            console.log('Getting mailbox settings...');
            results = await Promise.all(
                results.map(async (user: Map<string, string>) => {
                    const mailboxSettings = await this.getMailboxSettings(user.get('id') as string)
                    mappingMailboxSettings(user, mailboxSettings)
                    return user
                })
            )
            for (let user of results) {
                yield user;
            }
            console.log('New Iteration...');
            // Reset count 
            count = 0;
            results = []
            await pageIterator.resume();
        }
        console.log('Getting mailbox settings from last batch...');
        results = await Promise.all(
            results.map(async (user: Map<string, string>) => {
                const mailboxSettings = await this.getMailboxSettings(user.get('id') as string)
                mappingMailboxSettings(user, mailboxSettings)
                return user
            })
        )
        for (let user of results) {
            yield user;
        }
    }

    async getAccount(identity: string): Promise<Map<string, string>> {
        // let result: Map<string, string> = new Map()

        const result = await this.client
            .api('/users/' + identity)
            .top(1)
            .select([
                'id',
                'userPrincipalName',
                'mail',
                'displayName',
                'firstName',
                'lastName',
                'lastPasswordChangeDateTime',
            ])
            .get()
            .then((user) => {
                return mappingUser(user)
            })
            .catch((err) => {
                throw new ConnectorError('Unable to connect ' + err)
            })

        const mailboxSettings = await this.getMailboxSettings(identity as string)
        mappingMailboxSettings(result, mailboxSettings)
        return result
    }

    async testConnection(): Promise<any> {
        await this.client
            .api('/users')
            .top(1)
            .select(['id', 'userPrincipalName'])
            .get()
            .then(() => { }) // do nothing
            .catch((err) => {
                throw new ConnectorError('Unable to connect:' + err)
            })
    }
}
