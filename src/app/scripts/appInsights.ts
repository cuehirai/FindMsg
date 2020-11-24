import { ApplicationInsights, IEventTelemetry, IExceptionTelemetry, IMetricTelemetry, ITraceTelemetry, Snippet } from '@microsoft/applicationinsights-web';
import { AppConfig } from '../../config/AppConfig';

// set this to enable/disable AI completely
const useApplicationInsights = true;

// Look here for docs
// https://docs.microsoft.com/en-us/azure/azure-monitor/app/javascript
// https://github.com/Microsoft/ApplicationInsights-JS/blob/master/API-reference.md

function initAI(): ApplicationInsights | null {
    if (useApplicationInsights) {
        const config: Snippet = {
            config: {
                // instrumentationKey: '0b53b004-005f-46be-8fc1-63b98d9654e4',
                // Webapps(kacomslabo)
                // instrumentationKey: '474f6161-5de0-419b-baa4-9d57b2abdb37',
                instrumentationKey: AppConfig.AppInsight.instrumentationKey,
                maxBatchInterval: 5000,
                maxAjaxCallsPerView: -1,
                disableDataLossAnalysis: false,
                enableUnhandledPromiseRejectionTracking: true,
                disableFetchTracking: false,
            },
        };

        const _ai = new ApplicationInsights(config);
        _ai.loadAppInsights();
        _ai.trackPageView();

        return _ai;
    } else {
        return null;
    }
}

const instance = initAI();

function custom(callback: (instance: ApplicationInsights) => void): void {
    instance && callback && callback(instance);
}

function trackException(exception: IExceptionTelemetry): void {
    instance?.trackException(exception);
}

function trackEvent(event: IEventTelemetry, customProperties?: { [key: string]: string; }): void {
    instance?.trackEvent(event, customProperties);
}

function trackMetric(metric: IMetricTelemetry, customProperties?: { [key: string]: string; }): void {
    instance?.trackMetric(metric, customProperties);
}

function trackTrace(trace: ITraceTelemetry, customProperties?: { [key: string]: string; }): void {
    instance?.trackTrace(trace, customProperties);
}

function startTrackEvent(name?: string | undefined): void {
    instance?.startTrackEvent(name);
}

function stopTrackEvent(name: string, properties?: { [key: string]: string; }, measurements?: { [key: string]: number; }): void {
    instance?.stopTrackEvent(name, properties, measurements);
}

function flushBuffer() {
    instance?.flush();
}

function setAuthenticatedUserContext(userId: string, accountId?: string) {
    instance?.setAuthenticatedUserContext(userId, accountId, true);
}


export const AI = Object.freeze({
    custom,
    trackException,
    trackEvent,
    trackMetric,
    trackTrace,
    startTrackEvent,
    stopTrackEvent,
    flushBuffer,
    setAuthenticatedUserContext,
});
