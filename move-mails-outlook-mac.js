#!/usr/bin/env osascript -l JavaScript

"use strict";
var outlook = Application("Microsoft Outlook");
outlook.strictPropertyScope = true;
outlook.strictCommandScope = true;

// Entry point
function run(argv) {
    if (argv.length < 2) {
        return "usage: move-mails-outlook-mac.js <email> <folder>";
    }
    var destAccount = getDestAccount(argv[0]);
    checkNull(destAccount, "Cannot found dest account: " + argv[0]);
    var destFolder = getDestFolder(destAccount, argv[1]);
    checkNull(destFolder, "Cannot found dest folder: " + argv[1]);

    var srcAccount = outlook.defaultAccount();
    var srcFolder = srcAccount.inbox();

    console.log("Archiving messages from " + srcAccount.name() + "/Inbox to " + destAccount.name() + "/" + destFolder.name());
    archiveFolder(srcFolder, destFolder);
}
function getParentFolderName(folder) {
    return folder.container().name() != null ? "/" + folder.container().name() : "";
}
function getSubFolderByName(folder, name) {
    return folder.mailFolders().find(function (subFolder) {
        return subFolder.name() == name;
    });
}
function makeMailFolder(parentFolder, folderName) {
    console.log(parentFolder.name());
    console.log(folderName)
    return outlook.make({
        new: parentFolder.class(), at: parentFolder, withProperties: { name: folderName }
    });
}
function getSubFolderByName(folder, name) {
    var subFolder = getSubFolderByName(folder, name);
    return subFolder != null ? subFolder : makeMailFolder(folder, name);
}
function msgCntInFolderForArchiving(folder) {
    return folder.messages().length - folder.unreadCount();
}
function moveMsgs(srcFolder, destFolder) {
    srcFolder.messages().forEach(function (msg) {
        if (msg.isRead()) {
            console.log(msg.subject());
            //outlook.move(msg, { to: destFolder });
        }
    });
}
function archiveFolder(srcFolder, destFolder) {
    var msgCnt = msgCntInFolderForArchiving(srcFolder);
    var parentName = getParentFolderName(srcFolder);
    console.log("Archiving: %s/%s (%d)", parentName, srcFolder.name(), msgCnt);
    moveMsgs(srcFolder, destFolder);
    var srcSubFolders = new Array();
    srcFolder.mailFolders().forEach(function (srcSubFolder) {
        var destSubFolder = getSubFolderByName(destFolder, srcSubFolder.name());
        archiveFolder(srcSubFolder, destSubFolder);
    });
}
function getDestAccount(email) {
    return outlook.exchangeAccounts().find(function (elem) {
        return elem.emailAddress() == email;
    });
}
function getDestFolder(destAccount, folderName) {
    return destAccount.mailFolders().find(function (elem) {
        return elem.name() == folderName;
    });
}
function checkNull(obj, errorDesc) {
    if (!obj) {
        throw errorDesc;
    }
}







