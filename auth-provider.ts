import axios from 'axios'
import qs from 'qs'
import ora from 'ora'
import 'isomorphic-fetch'
import Conf from 'conf'
import { AuthenticationProvider } from '@microsoft/microsoft-graph-client'

const config = new Conf()

type DeviceAuthorizationResponse = {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message: string;
}

type SuccessfulAuthenticationResponse = {
  token_type: 'Bearer';
  scope: string;
  expires_in: number;
  access_token: string;
  id_token?: string;
  refresh_token: string;
}

type UnsuccessfulAuthenticationResponse = {
  error: string;
  error_description: string;
  error_codes: number[];
  timestamp: Date;
  trace_id: string;
  correlation_id: string;
  error_uri: string;
}

type AuthenticationResponse =
  SuccessfulAuthenticationResponse | UnsuccessfulAuthenticationResponse

function isSuccessful (response: AuthenticationResponse):
  response is SuccessfulAuthenticationResponse {
  return (response as SuccessfulAuthenticationResponse).token_type !== undefined
}

type TokenData = {
  accessToken: string;
  expireDate: number;
  refreshToken: string;
}

export type AppData = {
  id: string;
  secret: string;
  tenant: string;
  scope: string;
}

export class AuthProvider implements AuthenticationProvider {
  private printMessage: (message: string) => unknown

  private data?: TokenData

  private app: AppData

  public constructor (
    appData: AppData,
    printMessage: (message: string) => unknown = console.log
  ) {
    this.app = appData
    this.printMessage = printMessage
    this.data = config.get('auth')
    if (this.data) {
      console.log('Loaded saved data')
    }
  }

  private async initDeviceCodeFlow (): Promise<DeviceAuthorizationResponse> {
    const url =
      `https://login.microsoftonline.com/${this.app.tenant}/oauth2/v2.0/devicecode`
    const { data } = await axios.get(url, {
      params: {
        client_id: this.app.id,
        scope: this.app.scope
      },
      validateStatus: status => status >= 200 && status < 500
    })
    return data
  }

  private async pollToken (code: string): Promise<AuthenticationResponse> {
    const url =
      `https://login.microsoftonline.com/${this.app.tenant}/oauth2/v2.0/token`
    const { data } = await axios.post(url, qs.stringify({
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: this.app.id,
      device_code: code
    }), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      validateStatus: status => status >= 200 && status < 500
    })
    return data
  }

  private async refreshTokenQuery (refreshToken: string): Promise<AuthenticationResponse> {
    const url =
      `https://login.microsoftonline.com/${this.app.tenant}/oauth2/v2.0/token`
    const { data } = await axios.post(url, qs.stringify({
      grant_type: 'refresh_token',
      client_id: this.app.id,
      scope: this.app.scope,
      refresh_token: refreshToken
    }), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      validateStatus: status => status >= 200 && status < 500
    })
    return data
  }

  private async requestToken (): Promise<TokenData> {
    const { message, device_code, interval } = await this.initDeviceCodeFlow()
    this.printMessage(message)
    const spinner = ora('Waiting for authorization...').start()
    let tokenData: AuthenticationResponse = await this.pollToken(device_code)
    while (!isSuccessful(tokenData) && tokenData.error === 'authorization_pending') {
      // eslint-disable-next-line @typescript-eslint/no-unused-vars
      await new Promise((resolve, reject) => {
        setTimeout(resolve, interval * 1000)
      })
      tokenData = await this.pollToken(device_code)
    }
    spinner.stop()
    spinner.clear()

    if (!isSuccessful(tokenData)) {
      throw new Error('Device code has expired. Please, try again.')
    }

    return {
      accessToken: tokenData.access_token,
      expireDate: Date.now() + tokenData.expires_in * 1000,
      refreshToken: tokenData.refresh_token
    }
  }

  private async refreshToken (refreshToken: string): Promise<TokenData> {
    const tokenData = await this.refreshTokenQuery(refreshToken)

    if (!isSuccessful(tokenData)) {
      throw new Error(tokenData.error_description)
    }

    return {
      accessToken: tokenData.access_token,
      expireDate: Date.now() + tokenData.expires_in * 1000,
      refreshToken: tokenData.refresh_token
    }
  }

  /** @override */
  public async getAccessToken (): Promise<string> {
    if (!this.data) {
      this.data = await this.requestToken()
      config.set('auth', this.data)
    } else if (this.data.expireDate <= Date.now()) {
      console.log('Token has expired')
      this.data = await this.refreshToken(this.data.refreshToken)
      config.set('auth', this.data)
    }
    return this.data.accessToken
  }
}
