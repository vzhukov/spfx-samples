import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { AadHttpClient } from "@microsoft/sp-http";
import { HttpClient } from "@microsoft/sp-http";
import { AadTokenProviderFactory, AadTokenProvider } from "@microsoft/sp-http"
import IWhoAmI from "../model/IWhoAmI";
import { jwtDecode, JwtPayload } from "jwt-decode";

export default class DataverseService implements IDataverseService {
    private _aadTokenProviderFactory: AadTokenProviderFactory;
    private _httpClient: HttpClient;
    private _instanceUrl: string;
    private _token?: string;
    private _decodedToken?: JwtPayload;
    private _provider: AadTokenProvider;

    // Service key to consume the service
    public static readonly serviceKey: ServiceKey<DataverseService> = ServiceKey.create<DataverseService>('VitalyZhukovDataverseService', DataverseService);

    public constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            // Initializing token provider and http client
            this._aadTokenProviderFactory = serviceScope.consume(AadTokenProviderFactory.serviceKey);
            this._httpClient = serviceScope.consume(HttpClient.serviceKey);
        });
    }

    public async whoAmI(): Promise<IWhoAmI> {
        const url = `${this._instanceUrl}/api/data/v9.2/WhoAmI`;
        return this.get<IWhoAmI>(url);
    }

    public async getEntityList(entityPluralName: string, fields: string[]): Promise<any> {
        const url = `${this._instanceUrl}/api/data/v9.2/${entityPluralName}?$select=${fields.join(",")}`;
        const data = await this.get<any>(url);
        return data?.value || [];
    }

    /**
     * Execute GET-requests to the Dataverse API
     * @param url Endpoint address to proceed the request
     * @returns Response in JSON format
     */
    public async get<T>(url: string): Promise<T> {
        // Getting token
        const token = await this.getToken();

        // Perform request
        const response = await this._httpClient.get(url, AadHttpClient.configurations.v1, { headers: { Authorization: `Bearer ${token}` } });

        // Cast response to JSON-object
        const data = await response.json();

        if (response.ok) {
            return data;
        }
        else {
            throw data;
        }
    }

    public set instanceUrl(value: string) {
        //URL validation
        //const re = new RegExp(/^https\:\/\/[a-zA-Z0-9]{1,61}\.crm[0-9]{0,2}\.dynamics\.com$`/);
        this._instanceUrl = value;
        this._token = undefined;
        this._decodedToken = undefined;
    }

    public get instanceUrl(): string {
        return this._instanceUrl;
    }

    /**
     * Checking if the stored token expired
     */
    private get tokenExpired(): boolean {
        if (!this._token
            || !this._decodedToken?.exp) {
            return true;
        }

        const dt = new Date().getTime();
        const exp = new Date(this._decodedToken.exp * 1000).getTime();
        return exp - dt < 10000;
    }

    /**
     * Retrieve token to interact with Dataverse API
     * @returns JWT
     */
    private async getToken(): Promise<string | undefined> {
        if (this.tokenExpired) {
            // Getting instance of the AAD Token Provider
            this._provider = this._provider || await this._aadTokenProviderFactory.getTokenProvider();

            // Retrieve OBO Token
            this._token = await this._provider.getToken(this._instanceUrl);

            // Decode JWT to check expiration timestamp
            this._decodedToken = jwtDecode<JwtPayload>(this._token);
        }

        return this._token;
    }
}

/**
 * Service Declaration
 */
export interface IDataverseService {

    /**
     * Execute WhoAmI request
     * @returns WhoAmI response
     * @link https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/whoamiresponse?view=dataverse-latest
     */
    whoAmI(): Promise<IWhoAmI>;

    /**
     * Retrieve list of entities from Dataverse
     * @param entityPluralName Logical collection name of the entity (accounts, leads, etc.)
     * @param fields Fields to retrieve
     */
    getEntityList(entityPluralName: string, fields: string[]): Promise<any>;

    /**
     * Settings Dataverse instance in format https://[instance].[region].dynamics.com
     */
    set instanceUrl(value: string);

    get instanceUrl(): string;
}