<%@ Page Language="C#" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<WebPartPages:AllowFraming runat="server" />
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>OCC Search Data Bridge</title>
</head>
<body style="margin:0;font:12px/1.4 Segoe UI,Arial,sans-serif;color:#555;">
<div style="padding:8px;">OCC Search data bridge. This page is loaded automatically by the Service Management dashboard and is not meant to be opened directly.</div>
<script>
(function () {
    'use strict';

    var SM_ORIGIN = 'http://sharedspaces:8086';
    var SP_SITE = 'http://sharedspaces:8081/sites/OCC';
    var SP_LIST = 'GLK OCC Form';
    var LIST_GUID = null;
    var LIST_TITLE = null;

    var SELECT_FIELDS = [
        'Id', 'Title', 'Account_x0020_Name', 'Vertical', 'TypeOfIssue', 'Total_x0020_OCC_x0020_Value',
        'OCC_x0020_Reference_x0020_No', 'Status', 'OCC_x0020_Status', 'Created', 'Posted_Date',
        'Handset', 'No_of_Handsets', 'SalesDirectorName', 'Subject', 'RejectedReason',
        'Business_x0020_Justification', 'Annual_x0020_Revenue', 'Source', 'Services', 'ProductFamily', 'ProductCategory',
        'Trouble_x0020_Ticket', 'Sales_x0020_Comments', 'Marketing_x0020_Comments',
        'Sales_x0020_Director_x0020_Appro', 'Sales_x0020_Director_x0020_Appro0',
        'Head_x0020_Director_x0020_Approv', 'Head_x0020_Dir_x0020_Approval_x0', 'Head_x0020_Director_x0020_Text',
        'Channel_x0020_Manager_x0020_Vali', 'Channel_x0020_Manager_x0020_Vali0', 'Channel_x0020_Manager_x0020_Text',
        'Channel_x0020_Director_x0020_App', 'Channel_x0020_Director_x0020_App1', 'Channel_x0020_Director_x0020_Tex',
        'Channel_x0020_Director_x0020_App0', 'Head_x0020_of_x0020_Bid_x0020__x', 'Bid_x0020_Director_x0020_Text',
        'field11', 'field12', 'Head_x0020_of_x0020_Gov_x0020_LE', 'Head_x0020_of_x0020_Gov_x0020_LE0',
        'Head_x0020_of_x0020_LE_x002c__x0', 'CCO_x0020_Approval', 'CCOTimestamp', 'CCOLabel'
    ];
    var EXPAND = ['From', 'Channel_x0020_Manager_x0020_Name'];
    var EXPAND_SELECT = ['From/Title', 'From/EMail', 'Channel_x0020_Manager_x0020_Name/Title'];
    var DETAIL_EXTRA = [
        'Mail_x0020_Status_x0020_SD', 'Mail_x0020_Status_x0020_HD', 'Mail_x0020_Status_x0020_CD',
        'Mail_x0020_Status_x0020_AM', 'Mail_x0020_Status_x0020_PD', 'Mail_x0020_Status',
        'Assigned_x0020_to_x0020_AM_x0020', 'other_x0020__x002d__x0020_Type_x', 'Item_x0020_Url',
        'CO_TL_Assigned_Date', 'GOv_x0020_LE_x0020_Head_x0020_Te'
    ];

    // Same-origin XHR — identical to occ.html spGet
    function spGet(url, cb) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', url, true);
        xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
        xhr.onreadystatechange = function () {
            if (xhr.readyState !== 4) return;
            if (xhr.status >= 200 && xhr.status < 300) {
                try { cb(null, JSON.parse(xhr.responseText)); }
                catch (e) { cb(new Error('Bad JSON from SharePoint'), null); }
            } else {
                cb(new Error('HTTP ' + xhr.status), null);
            }
        };
        xhr.onerror = function () { cb(new Error('Network error'), null); };
        xhr.send();
    }

    function listApi(suffix) {
        if (!suffix || suffix.charAt(0) !== '/') suffix = '/' + (suffix || '');
        if (LIST_GUID) return SP_SITE + "/_api/web/lists(guid'" + LIST_GUID + "')" + suffix;
        var title = (LIST_TITLE || SP_LIST).replace(/'/g, "''");
        return SP_SITE + "/_api/web/lists/getbytitle('" + title + "')" + suffix;
    }

    function resolveList(cb) {
        if (LIST_GUID || LIST_TITLE) return cb(null);
        var title = SP_LIST.replace(/'/g, "''");
        spGet(SP_SITE + "/_api/web/lists/getbytitle('" + title + "')?$select=Id,Title", function (err, data) {
            if (!err && data && data.d) {
                LIST_GUID = data.d.Id || data.d.ID || null;
                LIST_TITLE = data.d.Title || SP_LIST;
            } else {
                LIST_TITLE = SP_LIST;
            }
            cb(null);
        });
    }

    function loadAllPages(url, acc, cb) {
        spGet(url, function (err, data) {
            if (err) return cb(err, acc);
            acc = acc.concat((data.d && data.d.results) || []);
            var next = data.d && data.d.__next;
            if (next) return loadAllPages(next, acc, cb);
            cb(null, acc);
        });
    }

    function loadAllItems(cb) {
        resolveList(function () {
            var select = SELECT_FIELDS.concat(EXPAND_SELECT).join(',');
            var url = listApi('/items?$select=' + encodeURIComponent(select) +
                '&$expand=' + EXPAND.join(',') + '&$top=5000&$orderby=Id desc');
            loadAllPages(url, [], cb);
        });
    }

    function loadItemById(itemId, cb) {
        resolveList(function () {
            var full = SELECT_FIELDS.concat(EXPAND_SELECT).concat(DETAIL_EXTRA).join(',');
            var url = listApi('/items(' + itemId + ')?$select=' + encodeURIComponent(full) +
                '&$expand=' + EXPAND.join(','));
            spGet(url, function (err, data) {
                if (!err) return cb(null, data.d);
                var core = SELECT_FIELDS.concat(EXPAND_SELECT).join(',');
                var url2 = listApi('/items(' + itemId + ')?$select=' + encodeURIComponent(core) +
                    '&$expand=' + EXPAND.join(','));
                spGet(url2, function (err2, data2) {
                    if (err2) return cb(err2, null);
                    cb(null, data2.d);
                });
            });
        });
    }

    function reply(ev, payload) {
        if (!ev.source) return;
        try { ev.source.postMessage(payload, SM_ORIGIN); } catch (e) {}
    }

    window.addEventListener('message', function (ev) {
        if (ev.origin !== SM_ORIGIN) return;
        var msg = ev.data;
        if (!msg || msg.type !== 'occ-bridge') return;

        if (msg.action === 'loadAll') {
            loadAllItems(function (err, rows) {
                reply(ev, {
                    type: 'occ-bridge', action: 'loadAll', requestId: msg.requestId,
                    error: err ? err.message : null, result: rows || []
                });
            });
            return;
        }

        if (msg.action === 'getItem') {
            var itemId = msg.payload && msg.payload.itemId;
            if (!itemId) {
                reply(ev, { type: 'occ-bridge', action: 'getItem', requestId: msg.requestId, error: 'Missing itemId', result: null });
                return;
            }
            loadItemById(itemId, function (err, row) {
                reply(ev, {
                    type: 'occ-bridge', action: 'getItem', requestId: msg.requestId,
                    error: err ? err.message : null, result: row
                });
            });
        }
    });

    if (window.parent && window.parent !== window) {
        window.parent.postMessage({ type: 'occ-bridge', action: 'ready' }, SM_ORIGIN);
    }
})();
</script>
</body>
</html>
