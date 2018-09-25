import { BearerTokenFetchClient, FetchOptions } from "@pnp/common";
import { AadTokenProvider } from "@microsoft/sp-http";

export class PortalTokenFetchClient extends BearerTokenFetchClient {
    constructor(private tokenProvider: AadTokenProvider) {
        super(null);
    }

    public fetch(url: string, options: FetchOptions = {}): Promise<Response> {
        return this.tokenProvider.getToken('https://chiverton365dev.sharepoint.com')
            .then((accessToken: string) => {
                this.token = accessToken;
                return super.fetch(url, options);
            });
    }
}
