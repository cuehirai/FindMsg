export interface IAppConfig {
    AppInfo: {
        name: string;
        appName: string;
        logo: string;
        host: string;
    };
    AuthClient: {
        AppId: string;
    };
    AppInsight: {
        instrumentationKey: string;
    };
}

const config = (): IAppConfig => {
    const res: IAppConfig = {
        AppInfo: {
            name: process.env.PACKAGE_NAME ?? "FindMsg",
            appName: process.env.APP_NAME ?? "Message Finder",
            logo: process.env.LOGO ?? "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product-fluent/svg/teams_48x1.svg",
            host: process.env.HOSTNAME ?? "",
        },
        AuthClient: {
            AppId: process.env.APPLICATION_ID ?? "",
        },
        AppInsight: {
            instrumentationKey: process.env.APPINSIGHTS_INSTRUMENTATIONKEY ?? "",
        },
    };

    return res;
}

export const AppConfig = config();
