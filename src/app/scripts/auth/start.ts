import { client } from './client';

(function login() {
    const url = new URL(location.href);
    const params = JSON.parse(url.searchParams.get("params") ?? JSON.stringify(null));

    client.loginRedirect(params);
})();
