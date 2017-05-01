#!/usr/bin/env osascript -l JavaScript

/*
 * move-mails-outlook-mac.js
 * Copyright (C) 2017 Reggie Zhang <reggy.zhang@gmail.com>
 * Licensed under the terms of The GNU Lesser General Public License (LGPLv3):
 * http://www.opensource.org/licenses/lgpl-3.0.html
 * 
 */

"use strict";

var Outlook = (function () {
    var instance;
    function init() {
        return Application("Microsoft Outlook");
    }
    return {
        getInstance: function () {
            if (!instance) {
                instance = init();
            }
            return instance;
        }
    };
})();

function archiveFolder(srcAccountName, srcFolder, destAccountName, destFolder) {
    console.log("Archiving messages from %s/%s to %s/%s", srcAccountName, srcFolder.name(), destAccountName, destFolder.name());
    doArchiving(srcFolder, destFolder);
}

function getMailFolderClass() {
    return Outlook.getInstance().defaultAccount().inbox().class();
}

function getParentFolderName(folder) {
    var parentFolderName = "";
    while (folder.container().name() != null) {
        parentFolderName = "/" + folder.container().name() + parentFolderName;
        folder = folder.container();
    }
    return parentFolderName;
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
    return Outlook.getInstance().make({
        new: parentFolder.class(), at: parentFolder, withProperties: { name: folderName }
    });
}

function msgCntInFolderForArchiving(folder) {
    return folder.messages().length - folder.unreadCount();
}

function moveMsgs(srcFolder, destFolder) {
    srcFolder.messages().forEach(function (msg) {
        if (msg.isRead()) {
            Outlook.getInstance().move(msg, { to: destFolder });
        }
    });
}

function doArchiving(srcFolder, destFolder) {
    var msgCnt = msgCntInFolderForArchiving(srcFolder);
    var parentName = getParentFolderName(srcFolder);
    console.log("Archiving: %s/%s (%d)", parentName, srcFolder.name(), msgCnt);
    moveMsgs(srcFolder, destFolder);
    var srcSubFolders = new Array();
    srcFolder.mailFolders().forEach(function (srcSubFolder) {
        var destSubFolder = getSubFolderByName(destFolder, srcSubFolder.name());
        if (destSubFolder != null) {
            doArchiving(srcSubFolder, destSubFolder);
        } else {
            var msgCnt = msgCntInFolderForArchiving(srcSubFolder);
            console.log("Skipping : /%s/%s (%d)", srcFolder.name(), srcSubFolder.name(), msgCnt);
        }
    });
}

function getDestAccount(email) {
    var destAccount = findAccount(Outlook.getInstance().exchangeAccounts(), email)
    return destAccount != null ? destAccount :  findAccount(Outlook.getInstance().imapAccounts(), email);
}

function findAccount(accounts, email) {
    return accounts.find(function (elem) {
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
    return destFolder != null ? destFolder : Outlook.getInstance().make({
        new: destAccount.inbox().class(), at: destAccount, withProperties: { name: folderName }
    });
}

function checkNull(obj, errorDesc) {
    if (!obj) {
        throw errorDesc;
    }
}

// Entry point
function run(argv) {
    if (argv.length < 2) {
        return "usage: move-mails-outlook-mac.js <email> <folder>";
    }
    var destAccount = getDestAccount(argv[0]);
    checkNull(destAccount, "Cannot found dest account: " + argv[0]);
    var destFolder = getDestFolder(destAccount, argv[1]);

    var srcAccount = Outlook.getInstance().defaultAccount();
    var srcInboxFolder = srcAccount.inbox();
    var srcSentFolder = srcAccount.sentItems();

    archiveFolder(srcAccount.name(), srcInboxFolder, destAccount.name(), destFolder);
    archiveFolder(srcAccount.name(), srcSentFolder, destAccount.name(), destFolder);
}
