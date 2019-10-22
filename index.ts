/* eslint-disable @typescript-eslint/camelcase */
import axios from 'axios'
import qs from 'qs'
import inquirer from 'inquirer'
import ora from 'ora'
import { config } from 'dotenv'
import 'isomorphic-fetch'
import { MailFolder, Message, User } from '@microsoft/microsoft-graph-types'
import { Client, AuthenticationProvider } from '@microsoft/microsoft-graph-client'

config()

const app = {
  id: process.env.APP_ID,
  secret: process.env.APP_SECRET,
  tenant: process.env.TENANT || 'common',
  scope: 'offline_access user.read mail.read'
}

class AuthProvider implements AuthenticationProvider {
  private token: string | null = null;

  private expire: number | null = Date.now();

  public constructor () {
    // Read saved token
  }

  public setToken (token: string): void {
    this.token = token
  }

  public async initDeviceCodeFlow (): Promise<any> {
    const url = `https://login.microsoftonline.com/${app.tenant}/oauth2/v2.0/devicecode`
    const { data } = await axios.get(url, {
      params: {
        client_id: app.id,
        scope: app.scope
      },
      validateStatus: status => status >= 200 && status < 500
    })
    return data
  }

  public async pollToken (code: string): Promise<any> {
    const url = `https://login.microsoftonline.com/${app.tenant}/oauth2/v2.0/token`
    const { data } = await axios.post(url, qs.stringify({
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: app.id,
      device_code: code
    }), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      validateStatus: status => status >= 200 && status < 500
    })
    return data
  }

  private updateToken () {
    // TODO
  }

  /**
   * This method will get called before every request to the msgraph server
   * This should return a Promise that resolves to an accessToken
   * (in case of success) or rejects with error (in case of failure)
   * Basically this method will contain the implementation for getting and
   * refreshing accessTokens
   */
  public async getAccessToken (): Promise<string> {
    return this.token || ''
  }
}

(async (): Promise<void> => {
  const authProvider = new AuthProvider()

  const deviceData = await authProvider.initDeviceCodeFlow()
  console.log(deviceData.message)
  const spinner = ora('Waiting for authorization...').start()
  let tokenData = await authProvider.pollToken(deviceData.device_code)
  while (!tokenData.access_token) {
    await new Promise((resolve, reject) => { setTimeout(resolve, deviceData.interval * 1000) })
    tokenData = await authProvider.pollToken(deviceData.device_code)
  }
  spinner.stop()

  authProvider.setToken(tokenData.access_token)

  const client = Client.initWithMiddleware({ authProvider })

  const me: User = await client.api('/me').get()
  console.log(`Authorized as ${me.displayName} (${me.mail})`)

  const mailFolders: { value: MailFolder[] } = await client.api('/me/mailFolders').get()

  const { folderId } = await inquirer.prompt([
    {
      type: 'list',
      name: 'folderId',
      message: 'Please select the folder to forward:',
      choices: mailFolders.value.map(folder => ({
        name: `${folder.displayName} (${folder.unreadItemCount}/${folder.totalItemCount})`,
        value: folder.id
      }))
    }
  ])

  const mail: { value: Message[] } = await client.api(`/me/mailFolders/${folderId}/messages/delta?$top=10`).get()

  console.log(mail.value.map(m => m.subject))
})().catch(error => console.error(`Error: ${error}`))
