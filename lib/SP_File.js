var Promise = require('bluebird');

class SP_File {

    constructor(client) {
        this.client = client;
    }

    upload(relativeUrl, rawData, fileName) {
        var root = this.client.getRootDir();
        var url = `/_api/web/getFolderByServerRelativeUrl('${root}${relativeUrl}')/Files/Add(overwrite=true,url='${fileName}')`;
        return this.client.post(url, rawData);
    }

    getFilesByRelativeUrl(relativeUrl, options) {
        var root = this.client.getRootDir();
        var query = this.client.getOptionsQuery(options);

        var url = `/_api/web/getFolderByServerRelativeUrl('${root}/${relativeUrl}')/Files` + (query ? `?${query}` : '');

        return this.client.get(url);
    }

    getFileByRelativeUrl(relativeUrl, options) {
        var root = this.client.getRootDir();
        var query = this.client.getOptionsQuery(options);

        var url = `/_api/web/getFileByServerRelativeUrl('${root}/${relativeUrl}')` + (query ? `?${query}` : '');
        return this.client.get(url);
    }

    downloadyRelativeUrl(relativeUrl) {
        var root = this.client.getRootDir();
        if (relativeUrl.indexOf('_vti_history') != -1) {
            root = '';
        }

        var url = `/_api/web/getFileByServerRelativeUrl('${root}${relativeUrl}')/$value`;
        return this.client.get(url);
    }

    downloadVersionByRelativeUrl(relativeUrl, versionId) {
        var root = this.client.getRootDir();
        var url = `/_api/web/getFileByServerRelativeUrl('${root}${relativeUrl}')/Versions(${versionId})/$value`;
        return this.client.get(url);
    }

    deleteByRelativeUrl(relativeUrl) {
        var root = this.client.getRootDir();
        var url = `/_api/web/getFileByServerRelativeUrl('${root}${relativeUrl}')/`;
        return this.client.delete(url, {});
    }

    versionsByRelativeUrl(relativeUrl, options) {
        var root = this.client.getRootDir();
        var query = this.client.getOptionsQuery(options);
        var url = `/_api/web/getFileByServerRelativeUrl('${root}${relativeUrl}')/Versions/` + (query ? `?${query}` : '');
        return this.client.get(url);
    }

    renameByRelativeUrl(relativeUrl, newFileName) {
        return new Promise((resolve, reject) => {
            var root = this.client.getRootDir();
            return this.getFileByRelativeUrl(relativeUrl).then(folders => {
                var folder = folders[0];
                var type = folder.ListItemAllFields.__metadata.type;

                var url = `/_api/web/getFileByServerRelativeUrl('${root}/${relativeUrl}')/ListItemAllFields`;
                return this.client.update(url, {
                    Title: newFileName,
                    FileLeafRef: newFileName,
                    __metadata: {
                        "type": type
                    }
                }, function (err, data) {
                    if (err) {
                        return reject(err);
                    }
                    return resolve(data);
                });
            }).catch(err => {
                return reject(err);
            })
        });
    }

}

module.exports = SP_File;