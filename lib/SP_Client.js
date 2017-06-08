var Promise = require('bluebird');
var SP_Folder = require('./SP_Folder');
var _ = require('lodash');

class SP_Client {

    constructor(client) {
        this.postClient = _.cloneDeep(client);
        this.updateClient = _.cloneDeep(client);
        this.client = _.cloneDeep(client);

        this.web = client.web;
        this.host = client.host;
        this.rootDir = client.rootDir;
    }

    get(url) {
        if (url.startsWith(this.client.url)) {
            url = url.substr(this.client.url.length);
        }
        url = this.removeDupliciteSlashes(url);


        var doRequest = require('sharepointconnector/lib/util/doRequest')(this.client);
        return new Promise((resolve, reject) => {
            return doRequest(url, (err, data) => {
                if (err) {
                    var msg = (err.message && err.message.value ? err.message.value : err);
                    return reject(new Error(msg));
                }
                if (typeof data == 'undefined' || data == null) {
                    data = [];
                } else if (!Array.isArray(data)) {
                    data = [data];
                }
                return resolve(data);
            });
        });
    }

    post(url, data) {
        url = this.removeDupliciteSlashes(url);
        var doRequest = require('sharepointconnector/lib/util/doRequest')(this.postClient);

        var requestData = {
            method: 'POST',
            headers: {
                'X-RequestDigest': this.postClient.baseContext,
                'Content-Type': 'application/x-www-form-urlencoded; odata=verbose'
            },
            processData: false,
            url: url
        };
        if (data instanceof  Buffer){
            requestData.formData = {
                my_buffer :data
            }
        }else{
            requestData.body = data;
        }


        this.postClient.baseHTTPOptions.headers['Content-Type'] = 'application/x-www-form-urlencoded; odata=verbose';
        this.postClient.baseHTTPOptions.headers['Accept'] = 'application/json; odata=verbose';
        return new Promise((resolve, reject) => {
            return doRequest(requestData, (err, data) => {
                if (err) {
                    var msg = (err.message && err.message.value ? err.message.value : err);
                    return reject(new Error(msg));
                }
                return resolve(data);
            });
        });
    }

    delete(url) {
        url = this.removeDupliciteSlashes(url);
        var doRequest = require('sharepointconnector/lib/util/doRequest')(this.client);
        return new Promise((resolve, reject) => {
            return doRequest({
                json: {},
                method: 'POST',
                headers: {
                    'X-RequestDigest': this.client.baseContext,
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE'
                },
                url: url
            }, (err, data) => {
                if (err) {
                    var msg = (err.message && err.message.value ? err.message.value : err);
                    return reject(new Error(msg));
                }
                return resolve(data);
            });
        });
    }

    update(url, updated, cb) {
        url = this.removeDupliciteSlashes(url);
        var doRequest = require('sharepointconnector/lib/util/doRequest')(this.updateClient);

        this.updateClient.baseHTTPOptions.headers['Content-Type'] = 'application/json; odata=verbose';
        this.updateClient.baseHTTPOptions.headers['Accept'] = 'application/json; odata=verbose';
        return doRequest({
            json: updated,
            method: 'POST',
            headers: {
                'X-RequestDigest': this.updateClient.baseContext,
                'X-HTTP-Method': 'MERGE',
                "IF-MATCH": "*"
            },
            url: url
        }, function (err) {
            if (err) {
                return cb(err);
            }
            return cb(null, updated);
        });
    }

    removeDupliciteSlashes(str) {
        while (str.indexOf('//') != -1) {
            str = str.replace('//', '/');
        }

        var root = this.getRootDir();
        while (str.indexOf(root + root) != -1) {
            str = str.replace(root + root, root);
        }

        return str;
    }

    getRootDir() {
        return '/' + this.client.web + this.client.rootDir;
    }

    getOptionsQuery(options) {
        if (!options) {
            return '';
        }

        var out = [];
        if (options.expand) {
            options.expand = (Array.isArray(options.expand) ? options.expand : [options.expand]);
            out.push('$expand=' + options.expand.join(','));
        }

        if (options.top) {
            out.push('$top=' + options.top);
        }

        if (options.skip) {
            out.push('$skip=' + options.skip);
        }

        if (options.orderby) {
            out.push('$orderby=' + options.orderby + (options.sort ? ` ${options.sort}` : ''));
        }

        return out.join('&');
    }

}

module.exports = SP_Client;