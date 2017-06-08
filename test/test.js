var assert = require('assert');
var fs = require('fs');
var Sharepoint = require(__dirname + '/../index.js');
var Promise = require('bluebird');

var sps;

describe('Sharepoint', function () {
    describe('Config', function () {
        it('Configuration file ./config/config.js is set.', function (done) {
            assert.equal(fs.existsSync(__dirname + '/../config/config.js'), true);
            done();
        });
    });
    describe('Api', function () {
        this.timeout(25000);
        before(function (done) {
            var config = require(__dirname + '/../config/config.js');
            sps = new Sharepoint();
            sps.login(config).then(client => {
                done();
            }).catch(err => {
                done(err);
            });
        });
        describe('SP_Folder', function () {
            it('[GET] root directory folders', function (done) {
                sps.SP_Folder.getFoldersByRelativeUrl('/').then(folders => {
                    done();
                }).catch(err => {
                    done(err);
                })
            });

            it('[POST] create "_apiFolder_123" directory in root', function (done) {
                sps.SP_Folder.create('/', '_apiFolder_123').then(folders => {
                    done();
                }).catch(err => {
                    console.log(err);
                    done(err);
                });
            });

            it('[DELETE] delete "_apiFolder_123" directory in root', function (done) {
                sps.SP_Folder.deleteByRelativeUrl('_apiFolder_123').then(folders => {
                    done();
                }).catch(err => {
                    done(err);
                });
            });

            it('[DELETE] change folder name', function (done) {
                var oldName = '_apiFolder_234';
                var newName = '_apiFolder_345';

                sps.SP_Folder.create('/', oldName).then(folders => {
                    sps.SP_Folder.renameByRelativeUrl(oldName, newName).then(ignored => {
                        assert.equal(ignored.Title, newName);
                        sps.SP_Folder.deleteByRelativeUrl(newName).then(folders => {
                            done();
                        }).catch(err => {
                            done(err);
                        });
                    }).catch(err => {
                        console.log(err);
                        done(err);
                    })
                }).catch(err => {
                    done(err);
                });
            });

        });

        describe('SP_File', function () {
            var fileContent = "Řepa čeká na ředkvičku v pískovišti."; // some utf-8 content
            var fileName = '_apiFile_123.dat';

            it(`[POST] create "${fileName}" file.`, function (done) {
                sps.SP_File.upload('/', fileContent, fileName).then(ignored => {
                    done();
                }).catch(err => {
                    done(err);
                });
            });

            it(`[POST] download "${fileName}" file.`, function (done) {
                sps.SP_File.downloadyRelativeUrl('/' + fileName).then(content => {
                    console.log(content);
                    assert(fileContent == content);
                    done();
                }).catch(err => {
                    done(err);
                });
            });

            it(`[POST] delete "${fileName}" file.`, function (done) {
                sps.SP_File.deleteByRelativeUrl('/' + fileName).then(ignored => {
                    done();
                }).catch(err => {
                    done(err);
                });
            });

            it(`[GET] revision "${fileName}" file.`, function (done) {
                var versions = [0, 1, 2, 3, 4]; // 5 files = 1 original, 4 revisions

                new Promise.mapSeries(versions, version => {
                    return sps.SP_File.upload('/', fileContent + version, fileName).then(ignored => {
                        return true;
                    }).catch(err => {
                        done(err);
                    });
                }).then(ignored => {
                    sps.SP_File.versionsByRelativeUrl('/' + fileName).then(revisions => {
                        assert.equal(versions.length - 1, revisions.length); // one file is original
                        sps.SP_File.deleteByRelativeUrl('/' + fileName).then(ignored => {
                            done();
                        }).catch(err => {
                            done(err);
                        });
                    }).catch(err => {
                        done(err);
                    })
                }).catch(err => {
                    done(err);
                })
            });
        });

    });
});
