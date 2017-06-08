var Promise = require('bluebird');
var SP_Folder = require('./lib/SP_Folder');
var SP_File = require('./lib/SP_File');

class Sharepoint {

    constructor() {
        this.workingDirectory = '';
    }

    login(params) {
        var self = this;
        return new Promise((resolve, reject) => {
            var client = {};
            client.url = params.host + params.web;
            client.web = params.web;
            client.rootDir = params.rootDir;
            client.federatedAuthUrl = params.federatedAuthUrl;
            client.auth = {};
            client.siteContext = params.context || 'web';
            client.auth.username = params.username;
            client.auth.password = params.password;
            client.auth.workstation = params.workstation || '';
            client.auth.domain = params.domain || '';
            client.auth.type = params.type || 'basic';
            client.auth.custom = params.authenticator;
            client.verbose = (typeof params.verbose === 'boolean') ? params.verbose : false;
            client.fieldValuesAsText = (typeof params.fieldValuesAsText === 'boolean') ? params.fieldValuesAsText : false;
            client.filterFields = params.filterFields || null; // an array of objects [{"field": one, "value": 1}, {"field": two, "value": 2}]
            client.selectFields = params.selectFields || null;  //array ["one", "two"]
            client.expandFields = params.expandFields || null; //array ["one", "two"]
            client.orOperator = (typeof params.orOperator === 'boolean') ? params.orOperator : false;
            client.andOrOperator = (typeof params.andOrOperator === 'boolean') ? params.andOrOperator : false;
            client.FedAuth = params.FedAuth || null;
            client.rtFa = params.rtFa || null;
            client.$top = params.$top || "5000";

            // We later set up http requests with client options object, either basic auth (goes in 'auth'),
            // or NTLM (goes in an Authorization header)
            client.baseHTTPOptions = {
                headers: {
                    'Accept': 'application/json; odata=verbose',
                    'Content-Type': 'application/x-www-form-urlencoded; odata=verbose'
                },
                strictSSL: (typeof params.strictSSL === 'boolean') ? params.strictSSL : true,
                proxy: params.proxy || undefined
            };

            var auths = {
                'basic': require('sharepointconnector/lib/auth/basic'),
                'ntlm': require('sharepointconnector/lib/auth/ntlm'),
                'online': require('sharepointconnector/lib/auth/online'),
                'onlinesaml': require('sharepointconnector/lib/auth/onlinesaml')
            };
            var makeSharepointResponsesLessAwful = require('sharepointconnector/lib/util/odata');

            var type = client.auth.type;
            if (!type || (!client.auth.custom && !auths.hasOwnProperty(type))) {
                return reject('Unsupported auth type: ' + type);
            }

// Either use a built-in authenticator or a supplied custom function
// which exposes the expected functionality
            client.httpClient = auths[type] || client.auth.custom;

            var context = (!client.siteContext || client.siteContext === 'web') ? '' : '/' + client.siteContext,
                contextUrl = client.url + context + '/_api/contextinfo';
            var httpOpts = {
                url: contextUrl,
                method: 'post'
            };

            // Call the implementing HTTP client, then handle the response generically
            return client.httpClient(client, httpOpts, function (err, response, body) {
                if (err) {
                    return reject(err);
                }
                if (response.statusCode.toString()[0] !== '2') {
                    return reject('Unexpected status code ' + response.statusCode);
                }

                body = makeSharepointResponsesLessAwful(body);
                if (!body.GetContextWebInformation || !body.GetContextWebInformation.FormDigestValue) {
                    return reject({
                        message: 'Failed to load default context from response body',
                        body: body
                    });
                }
                client.baseContext = body.GetContextWebInformation.FormDigestValue;
                var SP_Client = require('./lib/SP_Client');
                self.spClient = new SP_Client(client);
                return resolve(self.spClient);
            });
        });
    }

    get SP_Folder() {
        return new SP_Folder(this.spClient);
    }

    get SP_File() {
        return new SP_File(this.spClient);
    }

}

module.exports = Sharepoint;