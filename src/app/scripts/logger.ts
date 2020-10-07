import { isUnknownObject } from './utils';
import { AI } from './appInsights';

export function info(...args: unknown[]): void {
    console.log(...args); // tslint:disable-line:no-console
    args.filter(isUnknownObject).forEach(console.dir);
}

export function warn(...args: unknown[]): void {
    console.warn(...args);
    args.filter(isUnknownObject).forEach(console.dir);
}

export function error(...args: unknown[]): void {
    console.error(...args);
    args.filter(isUnknownObject).forEach(console.dir);
}


/**
 * Decorator to measure function execution time
 * @param _target unused
 * @param _propertyKey unused
 * @param descriptor
 */
export function trace(_target: unknown, _propertyKey: string, descriptor: PropertyDescriptor): PropertyDescriptor {
    const orig: (...args: unknown[]) => unknown = descriptor.value;
    const name: string = descriptor.value.name;

    descriptor.value = function traceWrapper(...args: unknown[]) {
        console.group(name);
        const start = performance.now();
        try {
            const res = orig.apply(this, args);
            if (isUnknownObject(res) && typeof res.then === "function") {
                console.warn(`trace: ${name} returned a promise. measured execution time may be wrong.`)
            }
            return res;
        } finally {
            const end = performance.now();
            const dt = (end - start).toFixed(1);
            console.info(`${name} took ${dt}ms`); // tslint:disable-line:no-console
            console.groupEnd();
        }
    };

    return descriptor;
}


/**
 * Decorator to measure function execution time of promise returning functions
 * @param _target unused
 * @param _propertyKey unused
 * @param descriptor
 */
export function traceAsync(logToAI = false) {
    return function _traceAsync(_target: unknown, _propertyKey: string, descriptor: PropertyDescriptor): PropertyDescriptor {
        const orig: (...args: unknown[]) => unknown = descriptor.value;
        const name: string = descriptor.value.name;

        descriptor.value = async function traceWrapper(...args: unknown[]) {
            console.group(name);
            const start = performance.now();
            try {
                if (logToAI) AI.startTrackEvent(name);
                const res = await orig.apply(this, args);
                return res;
            } finally {
                if (logToAI) AI.stopTrackEvent(name)
                const end = performance.now();
                const dt = (end - start).toFixed(1);
                console.info(`${name} took ${dt}ms`); // tslint:disable-line:no-console
                console.groupEnd();
            }
        };

        return descriptor;
    }
}