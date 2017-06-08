var Promise = require('bluebird');

class SP_Folder {

    constructor(client) {
        this.client = client;
    }

    getFoldersByRelativeUrl(relativeUrl, options) {
        var root = this.client.getRootDir();
        var query = this.client.getOptionsQuery(options);

        var url = `/_api/web/getFolderByServerRelativeUrl('${root}/${relativeUrl}')/Folders` + (query ? `?${query}` : '');
        return this.client.get(url);
    }

    getFolderByRelativeUrl(relativeUrl, options) {
        var query = this.client.getOptionsQuery(options);

        var root = this.client.getRootDir();
        var url = `/_api/web/getFolderByServerRelativeUrl('${root}/${relativeUrl}')` + (query ? `?${query}` : '');
        return this.client.get(url);
    }

    deleteByRelativeUrl(relativeUrl) {
        var root = this.client.getRootDir();
        var url = `/_api/web/getFolderByServerRelativeUrl('${root}/${relativeUrl}/')/`;
        return this.client.delete(url, {});
    }

    renameByRelativeUrl(relativeUrl, newFolderName) {
        return new Promise((resolve, reject) => {
            var root = this.client.getRootDir();
            return this.getFolderByRelativeUrl(relativeUrl, {expand: 'ListItemAllFields'}).then(folders => {
                var folder = folders[0];
                var type = folder.ListItemAllFields.__metadata.type;

                var url = `/_api/web/getFolderByServerRelativeUrl('${root}/${relativeUrl}')/ListItemAllFields`;
                return this.client.update(url, {
                    Title: newFolderName,
                    FileLeafRef: newFolderName,
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

    create(relativeUrl, folderName) {
        var root = this.client.getRootDir();
        var url = `/_api/web/getFolderByServerRelativeUrl('${root}${relativeUrl}')/Folders/Add('${folderName}')`;
        return this.client.post(url, '');
    }


}

module.exports = SP_Folder;