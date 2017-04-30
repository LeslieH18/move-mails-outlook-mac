#!/usr/bin/env osascript -l JavaScript

"use strict";
var outlook = Application("Microsoft Outlook");
var AUTO_CREATE_FOLDER = true;
// Entry point
function run(argv) {
    if (argv.length < 2) {
        return "usage: move-mails-outlook-mac.js <email> <folder>";
    }
    var destAccount = getDestAccount(argv[0]);
    checkNull(destAccount, "Cannot found dest account: " + argv[0]);
    var destFolder = getDestFolder(destAccount, argv[1]);

    var srcAccount = outlook.defaultAccount();
    var srcFolder = srcAccount.inbox();

    console.log("Archiving messages from " + srcAccount.name() + "/Inbox to " + destAccount.name() + "/" + destFolder.name());
    archiveFolder(srcFolder, destFolder);
}

function getParentFolderName(folder) {
    return folder.container().name() != null ? "/" + folder.container().name() : "";
}

function findSubFolderByName(folder, name) {
    return folder.mailFolders().find(function (subFolder) {
        return subFolder.name() == name;
    });
}

function getSubFolderByName(folder, name) {
    var subFolder = findSubFolderByName(folder, name);
    return subFolder != null ? subFolder : makeMailFolder(folder, name);
}

function makeMailFolder(parentFolder, folderName) {
    return outlook.make({
        new: parentFolder.class(), at: parentFolder, withProperties: { name: folderName }
    });
}

function msgCntInFolderForArchiving(folder) {
    return folder.messages().length - folder.unreadCount();
}

function moveMsgs(srcFolder, destFolder) {
    srcFolder.messages().forEach(function (msg) {
        if (msg.isRead()) {
            outlook.move(msg, { to: destFolder });
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
        if (destSubFolder != null) {
            archiveFolder(srcSubFolder, destSubFolder);
        } else {
            var msgCnt = msgCntInFolderForArchiving(srcSubFolder);
            console.log("Skipping : /%s/%s (%d)", srcFolder.name(), srcSubFolder.name(), msgCnt);
        }
    });
}

function getDestAccount(email) {
    return outlook.exchangeAccounts().find(function (elem) {
        return elem.emailAddress() == email;
    });
}

function findDestFolder(destAccount, folderName) {
    return destAccount.mailFolders().find(function (elem) {
        return elem.name() == folderName;
    });
}

function getDestFolder(destAccount, folderName) {
    var destFolder = findDestFolder(destAccount, folderName);
    return destFolder != null ? destFolder : outlook.make({
            new: destAccount.inbox().class(), at: destAccount, withProperties: { name: folderName }
        });
}

function checkNull(obj, errorDesc) {
    if (!obj) {
        throw errorDesc;
    }
}
