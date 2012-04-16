/*
FSO对象
*/

var g_KMHelperFSO = null;

var g_HL_Pink = "#ff00ff";
var g_HL_Blue = "#99ccff";
var g_HL_Yellow = "#ffff00";
var g_HL_Orange = "#ff9900";
var g_HL_Green = "#00ff00";
var g_HL_Purple = "#cc99ff";
var g_HL_Red = "#ff0000";

var g_KMSelection = null;
var g_KMSmartTagDivID = "WizKMSmartTagDivID";
var g_KMFlashMenuDivID = "WizKMFlashMenuDivID";
var g_KMCommentWindowDivID = "WizKMCommentWindowDivID";
var g_KMCommentWindowTextDivID = "KMCommentText";
var g_KMMousePos = { x: 0, y: 0};

RegKMButton();

function KMGetFSOObject() {
    if (!g_KMHelperFSO) {
        g_KMHelperFSO = objApp.CreateActiveXObject("Scripting.FileSystemObject");
    }
    return g_KMHelperFSO;
}


////=================================================================================
////文档关键字高亮代码
var KMHighlighter = function(colors) {
    this.colors = colors;
    if (this.colors == null) {
        //默认颜色   
        this.colors = ['#ffff00,#000000', '#dae9d1,#000000', '#eabcf4,#000000',
                       '#c8e5ef,#000000', '#f3e3cb, #000000', '#e7cfe0,#000000',
                       '#c5d1f1,#000000', '#deeee4, #000000', '#b55ed2,#000000',
                       '#dcb7a0,#333333', '#7983ab,#000000', '#6894b5, #000000'];
    }
}

KMHighlighter.prototype.highlight = function(doc, node, keywords, callback) {
    if (!keywords || !node || !node.nodeType || node.nodeType != 1)
        return;

    keywords = this.parsewords(keywords);
    if (keywords == null)
        return;
    //
    var text = node.innerText;
    //
    for (var i = 0; i < keywords.length; i++) {
        if (-1 == text.indexOf(keywords[i].word))
            continue;
        this.colorword(doc, node, keywords[i], callback);
    }
}

var g_KMHighlightSpanName = "wizKMHighlighterSpan";
//
KMHighlighter.prototype.colorword = function(doc, node, keyword, callback) {
    if (node.childNodes == undefined)
        return false;
    //
    if (node.name == g_KMHighlightSpanName)
        return false;
    //
    for (var i = 0; i < node.childNodes.length; i++) {
        var childNode = node.childNodes[i];
        if (childNode.nodeType == 3) {
            //childNode is #text
            var re = null;
            try {
                re = new RegExp(keyword.word, 'i');
            }
            catch (err) {
                continue;
            }
            if (!re)
                continue;
            //
            if (childNode.data.search(re) == -1)
                continue;
            //
            re = new RegExp('(' + keyword.word + ')', 'i');
            var forkNode = doc.createElement('span');
            forkNode.innerHTML = childNode.data.replace(re, '<span name="' + g_KMHighlightSpanName + '" style="background-color:' + keyword.bgColor + ';color:' + keyword.foreColor + '; cursor:hand; border-bottom: 1px #00c dashed;">$1</span>');
            node.replaceChild(forkNode, childNode);
            //
            for (var i = 0; i < forkNode.childNodes.length; i++) {
                var elem = forkNode.childNodes[i];
                if (elem.name == g_KMHighlightSpanName) {
                    elem.attachEvent("onclick", callback);
                }
            }
            //
            return true;
        } else if (childNode.nodeType == 1) {
            //childNode is element
            if (this.colorword(doc, childNode, keyword, callback))
                return true;
        }
    }
    return false;
}

KMHighlighter.prototype.parsewords = function(keywords) {
    var results = [];
    for (var i = 0; i < keywords.length; i++) {
        var keyword = {};
        var color = this.colors[i % this.colors.length].split(',');
        keyword.word = keywords[i];
        keyword.bgColor = color[0];
        keyword.foreColor = color[1];
        results.push(keyword);
    }
    return results;
}

KMHighlighter.prototype.sort = function(list) {
    list.sort(function(e1, e2){
              return e1.length < e2.length;
             });
}

////=================================================================================
////添加知识管理按钮并且相应该按钮消息，显示一个下拉框，该下拉框内容是一个html文件
//
function RegKMButton() {
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var languageFileName = pluginPath + "plugin.ini";
    var buttonWizhelper = objApp.LoadStringFromFile(languageFileName, "buttonWizhelper"); //WIZ_HELPER
    var buttonNoteplus = objApp.LoadStringFromFile(languageFileName, "buttonNoteplus");
    var buttonTagsCloud = objApp.LoadStringFromFile(languageFileName, "buttonTagsCloud");

    objWindow.AddToolButton("document", "KMHelperButton", buttonWizhelper, "", "OnKMHelperButtonClicked");
    objWindow.AddToolButton("document", "KMNoteplusButton", buttonNoteplus, "", "OnNotetakingButtonClicked");
    objWindow.AddToolButton("main", "KMTagsCloudButton", buttonTagsCloud, "", "OnKMTagsCloudButtonClicked");
    objWindow.AddToolButton("main", "KMHelperExButton", "【Status】", "", "OnKMHelperExButtonClicked");
}

function OnKMHelperButtonClicked() {
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var helperFileName = pluginPath + "KMHelperEx.htm";
    var rect = objWindow.GetToolButtonRect("document", "KMHelperButton");
    var arr = rect.split(',');
    // left,top,right,bottom
    objWindow.ShowSelectorWindow(helperFileName, arr[0], arr[3], 300, 500, "");
}

function OnNotetakingButtonClicked() {
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var noteFileName = pluginPath + "Notetaking.htm";
    //
    var rect = objWindow.GetToolButtonRect("document", "KMNoteplusButton");
    var arr = rect.split(',');
    objWindow.ShowSelectorWindow(noteFileName, arr[0], arr[3], 300, 500, "");
}

function OnKMHelperExButtonClicked() {
    if (objWindow.CurrentDocument != null) {
        var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
        var helperFileName = pluginPath + "KMHelperEx.htm";
        var rect = objWindow.GetToolButtonRect("main", "KMHelperExButton");
        var arr = rect.split(',');
        // left,top,right,bottom
        objWindow.ShowSelectorWindow(helperFileName, arr[0], arr[3], 300, 500, "");
    }
}

function OnKMTagsCloudButtonClicked() {
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var cloudFileName = pluginPath + "tagscloud.htm";
    //
    var rect = objWindow.GetToolButtonRect("main", "KMTagsCloudButton");
    var arr = rect.split(',');
    objWindow.ShowSelectorWindow(cloudFileName, arr[0], arr[3], 600, 600, "");
}


/*
获得所有的标签名称数组
*/
function KMGetAllTagsNameArray() {
    var objKMTags = objDatabase.Tags;

    var ret = [];
    for (var i = 0; i < objKMTags.Count; i++) {
        ret.push(objKMTags.Item(i).Name);
    }
    //
    return ret;
}

/*
获得所有的关键字数组。
首先从数据库查询所有的关键字，然后分割成关键字数组，分隔符为,;，；
*/
function KMGetAllKeywordsArray() {
    var objRowset = objDatabase.SQLQuery("select distinct document_keywords from wiz_document", "");

    //objKMTags
    var objMap = {};
    while (!objRowset.EOF) {
        var text = objRowset.GetFieldValue(0);
        text = text.replace(";", ",");
        text = text.replace("，", ",");
        text = text.replace("；", ",");
        var arr = text.split(',');
        //
        for (var i = 0; i < arr.length; i++) {
            var keyword = arr[i];
            objMap[keyword] = keyword;
        }
        objRowset.MoveNext();
    }

    var ret = [];
    for (var keyword in objMap) {
        ret.push(keyword);
    }
    //
    return ret;
}
/*
同样，获得所有的作者
*/
function KMGetAllAuthorArray() {
    var objRowset = objDatabase.SQLQuery("select distinct document_author from wiz_document", "");

    //objKMTags
    var objMap = {};
    while (!objRowset.EOF) {
        var text = objRowset.GetFieldValue(0);
        text = text.replace(";", ",");
        text = text.replace("，", ",");
        text = text.replace("；", ",");
        var arr = text.split(',');
        //
        for (var i = 0; i < arr.length; i++) {
            var author = arr[i];
            objMap[author] = author;
        }
        objRowset.MoveNext();
    }

    var ret = [];
    for (var author in objMap) {
        ret.push(author);
    }
    //
    return ret;
}

/*
点击关键字，标签，或者作者的时候，显示一个下拉列表窗口，列出相关的文档
*/
function KMShowListWindow(type) {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return false;
    //
    var e = doc.parentWindow.event;
    if (!e)
        return false;
    //
    var elem = e.srcElement;
    if (!elem)
        return false;
    var text = elem.innerText;
    if (text == null || text == "")
        return false;
    //
    var pluginpath = objApp.GetPluginPathByScriptFileName("km.js");
    var url = pluginpath + "listdocuments.htm?type=" + type + "&text=" + text;
    objWindow.ShowSelectorWindow(url, e.screenX, e.screenY, 350, 200, "");
}

/*
标签
*/
function KMTagWordSpanOnClick() {
    KMShowListWindow("tag");
}

/*
关键字
*/
function KMKeywordSpanOnClick() {
    KMShowListWindow("keyword");
}

/*
作者
*/
function KMAuthorSpanOnClick() {
    KMShowListWindow("author");
}

function KMGetFlashMenuStyle1() {
    return "cursor:hand; padding:2; display: block; width: 168px; text-align: left; text-decoration: none; font-family:arial; font-size:9pt; color: #000000; BORDER: none; border: solid 1px #FFFFFF;";
}

function KMGetFlashMenuStyle2() {
    return "cursor:hand; padding:2; display: block; width: 168px; text-align: left; text-decoration: none; font-family:arial; font-size:9pt; color: #000000; BORDER: none; border: solid 1px #6100C1;background-color:#F0E1FF";
}

//
function KMStringTrim(str) {
    return str.replace(/(^\s*)|(\s*$)/g, "");
}

function KMSmartTagGetSelectionText() {
    if (!g_KMSelection)
        return "";
    try {
        var text = g_KMSelection.text;
        if (!text)
            return "";
        text = KMStringTrim(text);
        return text;
    }
    catch (err) {
        return "";
    }
}
function KMTextToSingleLine(str) {
    str = str.replace(/\r/g, " ");
    str = str.replace(/\n/g, " ");
    str = str.replace(/\t/g, " ");
    return str;
}


function KMFlashSelectionAs(type) {
    var text = KMSmartTagGetSelectionText();
    if (!text || text == "")
        return;
    //
    text = KMTextToSingleLine(text);
    if (text.length > 200)
        text = text.substr(0, 200);
    //
    var objDoc = objApp.Window.CurrentDocument;
    if (!objDoc)
        return;
    //
    if (type == "title") {
        objDoc.ChangeTitleAndFileName(text);
    } else if (type == "tag") {
        text = objDoc.TagsText + ";" + text;
        objDoc.TagsText = text;
    } else if (type == "keyword") {
        text = objDoc.Keywords + ";" + text;
        objDoc.Keywords = text;
    } else if (type == "author") {
        objDoc.Author = text;
    }
// @endware
// D20110228 Add @RIL tag function
    else if ( type == "@RIL") {
        text= objDoc.TagsText + ";" + "@RIL";
        objDoc.TagsText = text;
    }

}
function KMFlashSelectionAsTitle() {
    KMFlashSelectionAs("title");
}
function KMFlashSelectionAsTags() {
    KMFlashSelectionAs("tag");
}
// @endware
// D20110228
function KMFlashSelectionAsRIL() {
    KMFlashSelectionAs("@RIL");
}
function KMFlashSelectionAsKeywords() {
    KMFlashSelectionAs("keyword");
}
function KMFlashSelectionAsAuthor() {
    KMFlashSelectionAs("author");
}


function KMFlashMenuItemMouseOver() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    var win = doc.parentWindow;
    if (!win)
        return;
    var event = win.event;
    //对于生成事件的 Window 对象、Document 对象或 Element 对象的引用。
    var src = event.srcElement;
    src.style.cssText = KMGetFlashMenuStyle2();
}
function KMFlashMenuItemMouseOut() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    var win = doc.parentWindow;
    if (!win)
        return;
    var event = win.event;
    //
    var src = event.srcElement;
    src.style.cssText = KMGetFlashMenuStyle1();
}

function KMAddFlashMenuItem(doc, divParent, text, callback) {
    var item = doc.createElement("DIV");
    item.style.cssText = KMGetFlashMenuStyle1();
    item.innerText = text;
    item.attachEvent("onclick", callback);
    item.attachEvent("onmouseover", KMFlashMenuItemMouseOver);
    item.attachEvent("onmouseout", KMFlashMenuItemMouseOut);
    divParent.appendChild(item);
}
function KMAddFlashMenuItem2(doc, divParent, textName, callback) {
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var languageFileName = pluginPath + "plugin.ini";
    //
    KMAddFlashMenuItem(doc, divParent, objApp.LoadStringFromFile(languageFileName, textName), callback);
}

function KMGetSelection(doc) {
    try {
        var sel = doc.selection.createRange();
        return sel;
    }
    catch (err) {
        return null;
    }
}
function KMCloneSelection(doc) {
    var sel = KMGetSelection(doc);
    if (!sel)
        return null;
    //
    return sel;
    //return sel.duplicate();
}

function KMFindChildNodeById(node, id) {
    if (node.attributes && node.attributes['id'] && node.attributes['id'].nodeValue == id) {
        return node;
    } else {
        if (node.hasChildNodes()) {
            var found = null;
            for (var i = 0; i < node.childNodes.length; i++) {
                found = KMFindChildNodeById(node.childNodes[i], id);
                if (found) {
                    return found;
                }
            }
        }
    }
    return null;
}

function KMRemoveChildNode(parentNode, nodeid) {
    var node = KMFindChildNodeById(parentNode, nodeid);
    if (!node)
        return;
    node.removeNode(true);
}
/*
function KMRemoveChildNodeByName(node, nodename) {
    if (node.hasChildNodes()) {
        for (var i = node.childNodes.length - 1; i >= 0; i--) {
            var child = node.childNodes[i];
            if (child.hasChildNodes()) {
                KMRemoveChildNodeByName(child, nodename);
            }
            if (child.name == nodename) {
                child.removeNode(false);
            }
        }
    }
    if (node.name == nodename) {
        node.removeNode(false);
    }
}
*/
function KMRemoveChildNodeByName(node, tagname, nodename) {
    var coll = node.getElementsByTagName(tagname);
    for (var i = 0; i < coll.length; i++) {
        var child = coll[i];
        if (child.name == nodename) {
            child.removeNode(false);
        }
    }
}

function KMRemoveChildNodeByID(node, id) {
	while ((child = KMFindChildNodeById(node, id)) != null) {
		child.removeNode(false);
	}
}

function KMGetHtmlHeaderFromText(text) {
    var index = text.search(/\<head/i);
    if (-1 == index)
        return "";
    return text.substr(0, index);
}
function KMGetHtmlFooterFromText(text) {
    return "</html>";
}

function KMGetDocumentFilePath(doc) {
    if (!doc)
        return "";

    var path = objApp.GetHtmlDocumentPath(doc);
    return path;
}
function KMGetDocumentFileName(doc) {
    if (!doc)
        return "";

    var path = objApp.GetHtmlDocumentPath(doc);
    var url = doc.URL;
    //
    var indexBookmark = url.indexOf("#");
    if (-1 != indexBookmark) {
        url = url.substr(0, indexBookmark);
    }
    //
    //
    url = url.replace(/\\/g, "/");
    var index = url.lastIndexOf("/");
    if (index == -1)
        return "";
    return path + url.substr(index + 1, url.length - index - 1);
}

/*
// save document
function KMSaveToDocument() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    //
    var objDocument = objWindow.CurrentDocument;
    if (!objDocument)
        return;
    //
    var filename = KMGetDocumentFileName(doc);
    var oldtext = objCommon.LoadTextFromFile(filename);
    //
    var header = KMGetHtmlHeaderFromText(oldtext);
    var footer = KMGetHtmlFooterFromText(oldtext);
    //
    var docElem = doc.documentElement.cloneNode(true);
    //
    KMRemoveChildNodeByName(docElem, g_KMHighlightSpanName);
    KMRemoveChildNode(docElem, g_KMSmartTagDivID);
    KMRemoveChildNode(docElem, g_KMCommentWindowDivID);
    KMRemoveChildNode(docElem, g_KMFlashMenuDivID);
    //
    var html = header + docElem.outerHTML + footer;
    //
    var template = html.indexOf("<!--WizHtmlContentBegin-->") != -1 && html.indexOf("<!--WizHtmlContentEnd-->") != -1;
    //
    try {
        objDatabase.BeginUpdate();  //不要发送更改消息，避免刷新网页
        var flags = 0x20; //不要下载网络文件，加快保存速度
        if (template) {
            flags |= 0x08;
        }
        objDocument.UpdateDocument4(html, filename, flags);
    }
    catch (err) {
    }
    objDatabase.EndUpdate();
}
*/

function KMGetCurrentInlineComment(doc, objCommentWindow) {
    if (!objCommentWindow) {
        objCommentWindow = KMGetCommentWindow(doc, false);
    }
    if (!objCommentWindow)
        return null;
    var inlineCommentID = objCommentWindow.getAttribute("KMCommentCurrentID");
    if (inlineCommentID == null || inlineCommentID == "")
        return null;
    //
    var objInlineComment = doc.getElementById(inlineCommentID);
    if (!objInlineComment)
        return null;
    return objInlineComment;
}

var g_KMDocumentModifiedAttributeName = "wizKMDocumentModified";

function KMSetDocumentModified(doc) {
    if (!doc)
        return;
    var body = doc.body;
    if (!body)
        return;
    //
    body.setAttribute(g_KMDocumentModifiedAttributeName, "1", 0);
}
function KMCancelDocumentModified(doc) {
    if (!doc)
        return;
    //
    var body = doc.body;
    if (!body)
        return;
    //
    body.removeAttribute(g_KMDocumentModifiedAttributeName);
}
function KMIsDocumentModified(doc) {
    if (!doc)
        return false;
    //
    var body = doc.body;
    if (!body)
        return false;
    //
    return body.getAttribute(g_KMDocumentModifiedAttributeName, 0) == "1";
}

function KMDeleteComment() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    var objCommentWindow = KMGetCommentWindow(doc, false);
    if (!objCommentWindow)
        return;
    //
    var objInlineComment = KMGetCurrentInlineComment(doc, objCommentWindow);
    //
    if (objInlineComment) {
        objInlineComment.removeNode(false);
        KMSetDocumentModified(doc);
    }

    KMCloseCommentWindow();
}
function KMCopyComment() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    var objCommentText = doc.getElementById(g_KMCommentWindowTextDivID);
    if (!objCommentText)
        return;
    var text = objCommentText.innerText;
    //
    doc.parentWindow.clipboardData.setData('Text', text);
}
function KMCloseCommentWindow() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    var objCommentWindow = KMGetCommentWindow(doc, false);
    if (!objCommentWindow)
        return;
    //
    objCommentWindow.style.display = "none";
}

function KMGetParentNodeByID(node, id) {
    while (node) {
        if (node.id == id)
            return node;
        node = node.parentElement;
    }
}

function KMIsChildNodeOf(node, id) {
    return KMGetParentNodeByID(node, id) != null;
}

function KMIsInlineCommentSpan(elem) {
    if (!elem.id)
        return false;
    if (elem.id.indexOf("WizKMComment_") != 0)
        return false;
    return true;
}
function KMGetParentInlineComment(elem) {
    while (elem != null) {
        if (KMIsInlineCommentSpan(elem))
            return elem;
        //
        elem = elem.parentElement;
    }
    return null;
}
function KMSetCommentWindowCaption(doc, caption) {
    var commentCaption = doc.getElementById("KMCommentCaption");
    if (commentCaption) {
        commentCaption.innerText = caption;
    }
}

function KMElemInInlineComment(elem) {
    if (null != KMGetParentInlineComment(elem))
        return true;
    return false;
}
function KMElemInCommentWindow(elem) {
    if (KMIsChildNodeOf(elem, g_KMCommentWindowDivID))
        return true;
    return false;
}



function KMAutoCloseCommentWindowTimer() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    //
    var objCommentWindow = KMGetCommentWindow(doc, false);
    if (!objCommentWindow)
        return;
    //
    if (objCommentWindow.style.display != "none") {
        var arr = [doc.activeElement, doc.elementFromPoint(g_KMMousePos.x, g_KMMousePos.y)];
        for (var i = 0; i < arr.length; i++) {
            var elem = arr[i];
            if (!elem)
                continue;
            if (KMElemInInlineComment(elem))
                return;
            if (KMElemInCommentWindow(elem))
                return;
        }
        //
        //
        var objCommentText = doc.getElementById(g_KMCommentWindowTextDivID);
        if (!objCommentText)
            return;
        var commentNew = objCommentText.innerHTML;
        var commentOld = "";
        //
        var objInlineComment = KMGetCurrentInlineComment(doc, objCommentWindow);
        if (objInlineComment) {
            commentOld = objInlineComment.title;
        }
        //
        if (commentNew != commentOld) {
            KMSaveComment();
        } else {
            objCommentWindow.style.display = "none";
        }
    }
    //
    objWindow.RemoveTimer("KMAutoCloseCommentWindowTimer");
}
function KMAutoCloseComment() {
    objWindow.RemoveTimer("KMAutoCloseCommentWindowTimer");
    objWindow.AddTimer("KMAutoCloseCommentWindowTimer", 1000);
}

function KMShowComment() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    //
    var win = doc.parentWindow;
    if (!win)
        return;
    //
    var e = win.event;
    if (!e)
        return;
    //
    var elem = e.srcElement;
    if (!elem)
        return;
    //
    var objInlineComment = KMGetParentInlineComment(elem);
    if (!objInlineComment)
        return;
    //
    var comment = objInlineComment.title;
    //
    var objCommentWindow = KMGetCommentWindow(doc, true);
    if (!objCommentWindow)
        return;
    //
    objCommentWindow.setAttribute("KMCommentCurrentID", objInlineComment.id);
    //
    var scrollX = Math.max(doc.body.scrollLeft, doc.documentElement.scrollLeft);
    var scrollY = Math.max(doc.body.scrollTop, doc.documentElement.scrollTop);
    var x = e.clientX + scrollX;
    var y = e.clientY + scrollY;
    //
    var widthCommentWindow = parseInt(objCommentWindow.style.width);
    //
    if (x + widthCommentWindow > doc.body.clientWidth) {
        x = doc.body.clientWidth - widthCommentWindow;
    }
    if (x < 0) {
        x = 0;
    }
    //
    y += 20;
    //
    objCommentWindow.style.left = x + "px";
    objCommentWindow.style.top = y + "px";
    objCommentWindow.style.display = "";
    //
    var commentText = doc.getElementById(g_KMCommentWindowTextDivID);
    if (commentText) {
        commentText.contentEditable = true;
        commentText.innerHTML = "";
        commentText.innerHTML = comment;
    }
    //
    KMSetCommentWindowCaption(doc, objInlineComment.innerText);
    //
    KMAutoCloseComment();
}


function KMGetRandomInt() {
    var d = new Date();
    return d.getTime();
}

function KMOnCommentShow() {
    KMShowComment();
}

function KMCommentAttachEvents(obj) {
    obj.attachEvent("onmouseover", KMOnCommentShow);
}

//
function KMSaveComment() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    var objCommentWindow = KMGetCommentWindow(doc, false);
    if (!objCommentWindow)
        return;
    var inlineCommentID = objCommentWindow.getAttribute("KMCommentCurrentID");
    if (inlineCommentID == null)
        inlineCommentID = "";
    //
    var objCommentText = doc.getElementById(g_KMCommentWindowTextDivID);
    if (!objCommentText)
        return;
    var comment = objCommentText.innerHTML;
    //
    if (inlineCommentID == "") {
        if (!g_KMSelection)
            return;
        var html = g_KMSelection.htmlText;
        //
        var id = "WizKMComment_" + KMGetRandomInt();
        //
        var newHtml = "<span id=\"" + id + "\" style=\"border-bottom: 2px #ff0000 dashed\" >" + html + "</span>";
        try {
            g_KMSelection.pasteHTML(newHtml);
            //
            var objInlineComment = doc.getElementById(id);
            if (objInlineComment) {
                objInlineComment.title = comment;
                KMCommentAttachEvents(objInlineComment);
            }
        }
        catch (err) {
        }
    } else {
        var objInlineComment = doc.getElementById(inlineCommentID);
        if (!objInlineComment)
            return;
        //
        objInlineComment.title = comment;
    }

    KMSetDocumentModified(doc);
    KMCloseCommentWindow();
}


function KMGetCommentWindow(doc, create) {
    var div = doc.getElementById(g_KMCommentWindowDivID);
    if (div != null)
        return div;
    if (!create)
        return null;
    //
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var languageFileName = pluginPath + "plugin.ini";
    //
    var commentHtml = objCommon.LoadTextFromFile(pluginPath + "inlinecomment.htm");
    commentHtml = commentHtml.replace("strOK", objApp.LoadStringFromFile(languageFileName, "strOK"));
    commentHtml = commentHtml.replace("strCancel", objApp.LoadStringFromFile(languageFileName, "strCancel"));
    commentHtml = commentHtml.replace("strCopy", objApp.LoadStringFromFile(languageFileName, "strCopy"));
    commentHtml = commentHtml.replace("strDelete", objApp.LoadStringFromFile(languageFileName, "strDelete"));
    commentHtml = commentHtml.replace("strClose", objApp.LoadStringFromFile(languageFileName, "strClose"));
    //
    div = doc.createElement("DIV");
    div.style.cssText = "z-index:100000;position:absolute;display:none;background:#e8e8e8;width:360px;height:168px;text-align:left;border:1px solid #c0c0c0;padding:5px 2px 5px 5px;filter: alpha(opacity=90)";
    div.id = g_KMCommentWindowDivID;
    //
    div.innerHTML = commentHtml;
    //
    var ret = doc.body.appendChild(div);
    //
    var imgClose = doc.getElementById("KMCommentCloseImage");
    if (imgClose) {
        imgClose.src = pluginPath + "close.gif";
        imgClose.attachEvent("onclick", KMCloseCommentWindow);
    }
    //
    var buttonOK = doc.getElementById("KMCommentOKButton");
    if (buttonOK) {
        buttonOK.attachEvent("onclick", KMSaveComment);
    }
    var buttonCancel = doc.getElementById("KMCommentCancelButton");
    if (buttonCancel) {
        buttonCancel.attachEvent("onclick", KMCloseCommentWindow);
    }
    var buttonCopy = doc.getElementById("KMCommentCopyButton");
    if (buttonCopy) {
        buttonCopy.attachEvent("onclick", KMCopyComment);
    }
    var buttonDelete = doc.getElementById("KMCommentDeleteButton");
    if (buttonDelete) {
        buttonDelete.attachEvent("onclick", KMDeleteComment);
    }
    //
    return ret;
}




function KMGetFlashMenuWindow(doc, create) {
    var div = doc.getElementById(g_KMFlashMenuDivID);
    if (div != null)
        return div;
    if (!create)
        return null;
    //
    div = doc.createElement("DIV");
    div.style.cssText = "padding:0; margin:0; position:absolute; z-index:100001; width:200px; display:none; background-color:window;border: solid 1px #6100C1;filter: progid:DXImageTransform.Microsoft.Shadow(color=#999999,direction=135,strength=5);";
    div.id = g_KMFlashMenuDivID;
    //
    KMAddFlashMenuItem2(doc, div, "strSelAsTitle", KMFlashSelectionAsTitle);
    KMAddFlashMenuItem2(doc, div, "strSelAsTags", KMFlashSelectionAsTags);
    KMAddFlashMenuItem2(doc, div, "strSelAsKeywords", KMFlashSelectionAsKeywords);
    KMAddFlashMenuItem2(doc, div, "strSelAsAuthor", KMFlashSelectionAsAuthor);
    //
    return doc.body.appendChild(div);
}

function KMCloseFlashMenu(doc, menu) {
    if (!menu) {
        menu = KMGetFlashMenuWindow(doc, false);
    }
    if (!menu)
        return;
    if (menu.style.display == "none")
        return;
    menu.style.display = "none";
}
function KMCloseSmartTagWindow(doc, smarttag) {
    if (!smarttag) {
        smarttag = KMGetSmartTagWindow(doc, false);
    }
    if (!smarttag)
        return;
    if (smarttag.style.display == "none")
        return;
    smarttag.style.display = "none";
}


// -----------------------------------------------------------------------------


function KMSearchEngine() {
    var _str = KMSmartTagGetSelectionText();
    if (!_str || _str == "")
        return;
    _str = KMTextToSingleLine(_str);
    WizExplorerApp.Window.ViewHtml('http://www.google.com.hk/search?q=' + _str, true);
}


function KMOnSmartTagFlashClick() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    //
    var objSmartTag = KMGetSmartTagWindow(doc, false);
    if (!objSmartTag)
        return;
    //
    var divMenu = KMGetFlashMenuWindow(doc, true);
    if (!divMenu)
        return;
    //
    KMCloseCommentWindow();
    KMCloseLinkToWindow(doc, null);
    KMCloseEditLinkToWindow(doc, null);
    KMCloseContentWindow();
    //
    divMenu.style.left = objSmartTag.offsetLeft + "px";
    divMenu.style.top = (objSmartTag.offsetTop + objSmartTag.offsetHeight) + "px";
    //
    divMenu.style.display = "";
}

function KMOnSmartTagCommentClick() {
    if (!g_KMSelection)
        return;
    //
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    //
    var objSmartTag = KMGetSmartTagWindow(doc, false);
    if (!objSmartTag)
        return;
    //
    var objCommentWindow = KMGetCommentWindow(doc, true);
    if (!objCommentWindow)
        return;
    //
    KMCloseFlashMenu(doc, null);
    KMCloseSearchWordWindow(doc, null);
    KMCloseLinkToWindow(doc, null);
    KMCloseContentWindow();
    KMCloseEditLinkToWindow(doc, null);
    //
    objCommentWindow.setAttribute("KMCommentCurrentID", "");
    //
    var x = objSmartTag.offsetLeft;
    var y = objSmartTag.offsetTop + objSmartTag.offsetHeight;
    //
    var widthCommentWindow = parseInt(objCommentWindow.style.width);
    //
    if (x + widthCommentWindow > doc.body.clientWidth) {
        x = doc.body.clientWidth - widthCommentWindow;
    }
    if (x < 0) {
        x = 0;
    }
    objCommentWindow.style.left = x + "px";
    objCommentWindow.style.top = y + "px";
    objCommentWindow.style.display = "";
    //
    //
    var commentText = doc.getElementById(g_KMCommentWindowTextDivID);
    if (commentText) {
        commentText.contentEditable = true;
        commentText.innerHTML = "";
    }
    //
    KMSetCommentWindowCaption(doc, g_KMSelection.text);
    //
    KMAutoCloseComment();
}

function KMChangeSelectionBackColor(color) {
    if (!g_KMSelection)
        return;
    //
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    //
    /*
    //@endware
    if  (color == "#ff0000") {
        var span_id= "hl_mark_" + KMGetRandomInt();
        //g_KMSelection.execCommand("FormatBlock",false,"<span " + span_id + ">");
        g_KMSelection.elements[0].setAttribute("wiz_mark",span_id);
        g_KMSelection.execCommand("BackColor", false, color);
        
    } else {	
    */
    g_KMSelection.execCommand("BackColor", false, color);
    //}
    //
    KMSetDocumentModified(doc);
    //
    KMCloseCommentWindow();
    KMCloseFlashMenu(doc, null);
    KMCloseSearchWordWindow(doc, null);
    KMCloseLinkToWindow(doc, null);
    KMCloseEditLinkToWindow(doc, null);
    KMCloseContentWindow();
    KMCloseSmartTagWindow(doc, null);
}

function KMBookmarkAddText() {
    var ForAppending = 8;
    var objDB = objApp.Database;
    var objDoc = objApp.Window.CurrentDocument;
    var dbPath = objDB.DatabasePath;
    var bookmark = dbPath + "bookmark.txt"

    var objFSO = objApp.CreateActiveXObject("Scripting.FileSystemObject");
    //
    try {
        if (!objFSO.FileExists(bookmark))
            var objFile = objFSO.CreateTextFile(bookmark);
        else 
            var objFile = objFSO.GetFile(bookmark);
    }
    catch (err) {
        //
    }

    var bmFile = objFile.OpenAsTextStream(ForAppending, false);
    var today = new Date();
    var nameBookmark = g_KMSelection.text.replace(/\"|'|`/g,"");
//    pToday = today.getYear() + "-" + (today.getMonth() +1) + "-" + today.getDate() + " " + today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
    pToday = today.getYear() + "-" + (today.getMonth() +1) + "-" + today.getDate();
    bmFile.WriteLine("");
    bmFile.WriteLine("<Row>");
    bmFile.WriteLine("\t<Cell><Data ss:Type=\"Date\">" + pToday + "</Data></Cell>");
    bmFile.WriteLine("\t<Cell><Data ss:Type=\"String\">" + nameBookmark + "</Data></Cell>");
    bmFile.WriteLine("\t<Cell><Data ss:Type=\"String\">" + objDoc.Title + "</Data></Cell>");
    bmFile.WriteLine("</Row>");
    bmFile.WriteLine("");
    bmFile.Close();
}


function KMBookmarkAdd() {
    if (!g_KMSelection) {
        return;
    }

    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;

    KMBookmarkAddText();
    var nameBookmark = g_KMSelection.text.replace(/\"|'|`/g,"");
    g_KMSelection.execCommand("CreateBookmark",false,nameBookmark);

    KMSetDocumentModified(doc);
    //
    KMCloseCommentWindow();
    KMCloseFlashMenu(doc, null);
    KMCloseSearchWordWindow(doc, null);
    KMCloseLinkToWindow(doc, null);
    KMCloseEditLinkToWindow(doc, null);
    KMCloseContentWindow();
    KMCloseSmartTagWindow(doc, null);

}

function KMBookmarkRemove() {
}


// FF0000(ff7f27) Red
function KMOnSmartTagRedClick() {
    KMChangeSelectionBackColor(g_HL_Red);
}

//FF00FF Pink
function KMOnSmartTagPinkClick() {
    KMChangeSelectionBackColor(g_HL_Pink);
}

//99CCFF Blue
function KMOnSmartTagBlueClick() {
    KMChangeSelectionBackColor(g_HL_Blue);
}

//FFFF00 Yellow
function KMOnSmartTagYellowClick() {
    KMChangeSelectionBackColor(g_HL_Yellow);
}

//00FF00(b5e61e) Green
function KMOnSmartTagGreenClick() {
    KMChangeSelectionBackColor(g_HL_Green);
}

// FF9900 Orange
function KMOnSmartTagOrangeClick() {
    KMChangeSelectionBackColor(g_HL_Orange);
}

// CC99FF Purple
function KMOnSmartTagPurpleClick() {
    KMChangeSelectionBackColor(g_HL_Purple);
}

function KMOnSearchEngineClick() {
    KMSearchEngine("");
}

function KMOnBookmarkClick() {
    KMBookmarkAdd();
}


function KMOnSmartTagEraserClick() {
    KMChangeSelectionBackColor("");
}

function KMAddSmartTagButton(pluginPath, doc, objSmartWindow, imgfilename, buttonname, id, callback) {
    var img = doc.createElement("IMG");
    img.border = 0;
    img.src = pluginPath + imgfilename;
    img.title = objApp.LoadStringFromFile(pluginPath + "plugin.ini", buttonname);
    img.id = id;
    img.style.cssText = "padding:0; margin:0; cursor:hand; border-style:none;";
    img.attachEvent("onclick", callback);
    //
    objSmartWindow.appendChild(img);
}

function KMGetSmartTagWindow(doc, create) {
    var objSmartTagWindow = doc.getElementById(g_KMSmartTagDivID);
    if (objSmartTagWindow != null)
        return objSmartTagWindow;
    if (!create)
        return null;
    //
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    //
    objSmartTagWindow = doc.createElement("DIV");
    objSmartTagWindow.style.cssText = "padding:0; margin:0; cursor:hand; position:absolute; z-index:100000; width:200px; height:48px; background-color:transparent; display:none;";
    objSmartTagWindow.id = g_KMSmartTagDivID;
    // Row 1: Highlighter
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_pink.png", "strMarkPink", "KMSmartTagPinkImg", KMOnSmartTagPinkClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_blue.png", "strMarkBlue", "KMSmartTagBlueImg", KMOnSmartTagBlueClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_yellow.png", "strMarkYellow", "KMSmartTagYellowImg", KMOnSmartTagYellowClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_orange.png", "strMarkOrange", "KMSmartTagOrangeImg", KMOnSmartTagOrangeClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_green.png", "strMarkGreen", "KMSmartTagGreenImg", KMOnSmartTagGreenClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_purple.png", "strMarkPurple", "KMSmartTagPurpleImg", KMOnSmartTagPurpleClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_red.png", "strMarkRed", "KMSmartTagRedImg", KMOnSmartTagRedClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_eraser.png", "strEraser", "KMSmartTagEraserImg", KMOnSmartTagEraserClick);
    // Row 2: Others
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_bookmark.png", "strBookmark", "KMSmartTagBookmarkImg", KMOnBookmarkClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_contents.png", "strContents","KMSetAsContent", KMOnSetAsContentClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_google.png", "strSeIcon", "KMSmartTagSeIconImg", KMOnSearchEngineClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_dictionary.png", "strSWIcon", "KMSmartTagSeWordImg", KMOnSearchWordClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_comment.png", "strComment", "KMSmartTagCommentImg", KMOnSmartTagCommentClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_linkto.png", "strLinkTo","KMSmartTagLinkToImg", KMOnLinkToClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_flash.png", "strKM", "KMSmartTagFlashImg", KMOnSmartTagFlashClick);
    KMAddSmartTagButton(pluginPath, doc, objSmartTagWindow, "tm_readitlater.png", "strReadItLater","KMSmartTagRILImg", KMFlashSelectionAsRIL);
    //
    return doc.body.appendChild(objSmartTagWindow);
}

function KMOnMouseMove() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    var win = doc.parentWindow;
    if (!win)
        return;
    var event = win.event;
    //
    g_KMMousePos.x = event.clientX;
    g_KMMousePos.y = event.clientY;
}
function KMIsSelected(sel) {
    if (!sel)
        return false;
    if (!sel.text)
        return false;
    var text = sel.text;
    if (text.length == 0)
        return false;
    text = KMStringTrim(text);
    if (text.length == 0)
        return false;
    return true;
}
function KMOnMouseUp() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    var win = doc.parentWindow;
    if (!win)
        return;
    var event = win.event;
    //
    // if (event.button != 1)
    //     return;
    //
    var src = event.srcElement;
    if (src.id == g_KMSmartTagDivID)
        return;
    if (src.parentElement.id == g_KMSmartTagDivID)
        return;
    if (src.parentElement.id == g_KMCommentWindowTextDivID)
        return;
    //
    var sel = KMGetSelection(doc);
    var objDatabase = objApp.Database;
    var isAltDownNeeded = objDatabase.Meta("wizhelp_parm","ALTKEY_FLAG");
    if (KMIsSelected(sel) && (event.altKey || isAltDownNeeded!="1") ) {
        //
        g_KMSelection = sel;
        //
        var scrollX = Math.max(doc.body.scrollLeft, doc.documentElement.scrollLeft);
        var scrollY = Math.max(doc.body.scrollTop, doc.documentElement.scrollTop);
        if (event.button == 1) {
            var x = event.clientX + scrollX;
            var y = event.clientY + scrollY;
            y += 20;
        }
        else {
            var markerTextCharEntity = "&#xfeff;";
            var markerId = "sel_" + new Date().getTime() + "_" + Math.random().toString().substr(2);
            var range = sel.duplicate();
            range.collapse(false);
            range.pasteHTML('<span id="' + markerId + '"  style="POSITION:absolute;">' + markerTextCharEntity + '</span>');
            var markerEl = doc.getElementById(markerId);
            var x = markerEl.offsetLeft; 
            var y = markerEl.offsetTop;
            y += 30;
            markerEl.parentNode.removeChild(markerEl);
        }
        if (x > (scrollX+doc.body.clientWidth-200)) {
            x = scrollX+doc.body.clientWidth-200;
        }
        if (y > (scrollY+doc.body.clientHeight-48)) {
            y -= 80;
        }
        var objSmartTag = KMGetSmartTagWindow(doc, true);
       //
        objSmartTag.style.display = "";
        objSmartTag.style.left = x + "px";
        objSmartTag.style.top = y + "px";
    }
    else {
        var objSmartTag = KMGetSmartTagWindow(doc, false);
        if (objSmartTag != null) {
            objSmartTag.style.display = "none";
        }
    }
    //
    KMCloseFlashMenu(doc, null);
}

function KMKeyMon() {
//	objDB.Meta("TEST","MSG_311") = "t1";
    var doc = objWindow.CurrentDocumentHtmlDocument;
    var objDoc = objWindow.CurrentDocument;
    if (!doc)
        return;
    var win = doc.parentWindow;
    if (!win)
        return;
    var event = win.event;

    var src = event.srcElement;
    if (src.id == g_KMSmartTagDivID)
        return;
    if (src.parentElement.id == g_KMSmartTagDivID)
        return;
    if (src.parentElement.id == g_KMCommentWindowTextDivID)
        return;

    if (event.altKey) {
        switch (event.keyCode) {
        case 18: //alt
            KMOnMouseUp();
            break;
        case 49: //alt+1
            var sel = KMGetSelection(doc);
            if (KMIsSelected(sel)) {
                g_KMSelection = sel;
                g_KMSelection.execCommand("BackColor", false, g_HL_Pink);
                KMSetDocumentModified(doc);
            }
            break;
        case 50: //alt+2
            var sel = KMGetSelection(doc);
            if (KMIsSelected(sel)) {
                g_KMSelection = sel;
                g_KMSelection.execCommand("BackColor", false, g_HL_Blue);
                KMSetDocumentModified(doc);
            }
            break;
        case 51: //alt+3
            var sel = KMGetSelection(doc);
            if (KMIsSelected(sel)) {
                g_KMSelection = sel;
                g_KMSelection.execCommand("BackColor", false, g_HL_Yellow);
                KMSetDocumentModified(doc);
            }
            break;
        case 52: //alt+4
            var sel = KMGetSelection(doc);
            if (KMIsSelected(sel)) {
                g_KMSelection = sel;
                g_KMSelection.execCommand("BackColor", false, g_HL_Orange);
                KMSetDocumentModified(doc);
            }
            break;
        case 53: //alt+5
            var sel = KMGetSelection(doc);
            if (KMIsSelected(sel)) {
                g_KMSelection = sel;
                g_KMSelection.execCommand("BackColor", false, g_HL_Green);
                KMSetDocumentModified(doc);
            }
            break;
        case 54: //alt+6
            var sel = KMGetSelection(doc);
            if (KMIsSelected(sel)) {
                g_KMSelection = sel;
                g_KMSelection.execCommand("BackColor", false, g_HL_Purple);
                KMSetDocumentModified(doc);
            }
            break;
        case 55: //alt+7
            var sel = KMGetSelection(doc);
            if (KMIsSelected(sel)) {
                g_KMSelection = sel;
                g_KMSelection.execCommand("BackColor", false, g_HL_Red);
                KMSetDocumentModified(doc);
            }
            break;
        case 48: case 192: //alt+0, alt+` clean highligth
            var sel = KMGetSelection(doc);
            if (KMIsSelected(sel)) {
                g_KMSelection = sel;
                g_KMSelection.execCommand("BackColor", false, "");
                KMSetDocumentModified(doc);
            }
            break;
        case 82: //alt+r show flash menu
            KMFlashSelectionAsRIL
            break;
        case 83: //alt+s save current changes
            KMSaveDocument(doc, objDoc);
            break;
        case 84: //alt+t add tag for selecttion
            KMFlashSelectionAsTags();
            break;
        default:
            break;
        }
    }
}


function KMAttachCommentEvents(doc) {
    var arr = doc.getElementsByTagName("SPAN");
    for (var i = 0; i < arr.length; i++) {
        var elem = arr[i];
        if (!KMIsInlineCommentSpan(elem))
            continue;
        //
        KMCommentAttachEvents(elem);
    }
}


function KMGetFileSize(filename) {
    var objFSO = KMGetFSOObject();
    //
    try {
        if (!objFSO.FileExists(filename))
            return 0;
        var objFile = objFSO.GetFile(filename);
        return objFile.Size;
    }
    catch (err) {
    }
    return 0;
}


/*
文档显示完成的时候，高亮显示标签，关键字或者作者
*/
function KMOnHtmlDocumentComplete(doc) {
    var filename = KMGetDocumentFileName(doc);
    var filesize = KMGetFileSize(filename);
    //objDB.Meta("TEST","MSG_200")= doc.ParamValue('DOC_READ_COUNT');

    //
    try {
        var url = doc.URL;
        if (url == null)
            return;
        url = url.toLowerCase();
        if (url.indexOf("plugins") != -1)
            return;
    }
    catch (err) {
        return;
    }
    
    
    //@endware
    //	if (filesize != 0 && filesize < 1024 * 500) {
    if ( objDB.Meta("wizhelp_parm","keyword_flag") == "1"  ) {
        var hl = new KMHighlighter();
        hl.highlight(doc, doc.body, KMGetAllTagsNameArray(), KMTagWordSpanOnClick);
        hl.highlight(doc, doc.body, KMGetAllKeywordsArray(), KMKeywordSpanOnClick);
        hl.highlight(doc, doc.body, KMGetAllAuthorArray(), KMAuthorSpanOnClick);
    }
    //@endware
    // smarttag flag
    if ( objDB.Meta("wizhelp_parm","smarttag_flag") == "1") {
        doc.attachEvent("onmouseup", KMOnMouseUp);
        doc.attachEvent("onmousemove", KMOnMouseMove);
    }

    doc.attachEvent("onkeydown",KMKeyMon);
    //

    KMAttachCommentEvents(doc);
    KMAttachLinkToEvents(doc);
    if (objWindow.CurrentDocument != null) {
        UpdateButtonStatus();
    }
}



/*

function KMOnWindowUnload() {
    var arr = [];
    for (var g in g_KMDocumentGUIDFileNameMap) {
        var filename = g_KMDocumentGUIDFileNameMap[g];
        if (!filename)
            continue;
        var guid = g.substr(1);
        guid = guid.replace(/\_/g, "-");
        try {
            var objDocument = objDatabase.DocumentFromGUID(guid);
            arr.push(objDocument);
        }
        catch (err) {
        }
    }
    //
    if (arr.length == 0)
        return;
    //
    try {
        var progress = objApp.CreateWizObject("WizKMControls.WizProgressWindow");
        progress.Title = "Saving document";
        progress.Text = objDocument.Title;
        progress.Show();
        progress.Max = arr.length;
        //
        for (var i = 0; i < arr.length; i++) {
            var objDocument = arr[i];
            var filename = KMGetDocumentModifiedFileName(objDocument);
            if (!filename)
                continue;
            //
            KMSaveToDocument(objDocument, filename);
            KMCancelDocumentModified(objDocument);
            //
            progress.Pos = i + 1;
        }
        progress.Hide();
        progress.Destroy();
    }
    catch (err) {
        WizAlert(err.message);
        //
    }
}
*/


function KMSaveDocumentCore(objHtmlDocument, objWizDocument) {
    if (!objWizDocument)
        return;
    if (!KMIsDocumentModified(objHtmlDocument))
        return;
    //
    KMCancelDocumentModified(objHtmlDocument);
    //
    var srcfilename = KMGetDocumentFileName(objHtmlDocument);
    var oldtext = objCommon.LoadTextFromFile(srcfilename);
    //
    var header = KMGetHtmlHeaderFromText(oldtext);
    var footer = KMGetHtmlFooterFromText(oldtext);
    //
    var docElem = objHtmlDocument.documentElement.cloneNode(true);
    //
    KMRemoveChildNodeByName(docElem, "span", g_KMHighlightSpanName);
	KMRemoveChildNodeByID(docElem, "wizkm_highlight");
    KMRemoveChildNode(docElem, g_KMSmartTagDivID);
    KMRemoveChildNode(docElem, g_KMCommentWindowDivID);
    KMRemoveChildNode(docElem, g_KMFlashMenuDivID);
    KMRemoveChildNode(docElem, "WizKMSearchWordDivID");
    KMRemoveChildNode(docElem, "WizKMEditLinkToDivID");
    KMRemoveChildNode(docElem, "WizKMLinkToFolderDivID");
    KMRemoveChildNode(docElem, "WizKMContentDivID");
    
    //
    var html = header + docElem.outerHTML + footer;
    //
    var template = html.indexOf("<!--WizHtmlContentBegin-->") != -1 && html.indexOf("<!--WizHtmlContentEnd-->") != -1;
    //
    try {
        objDatabase.BeginUpdate();  //不要发送更改消息，避免刷新网页
        var flags = 0x20; //不要下载网络文件，加快保存速度
        if (template) {
            flags |= 0x08;
        }
        objWizDocument.UpdateDocument4(html, srcfilename, flags);
    }
    catch (err) {
        WizAlert(err.message);
    }
    objDatabase.EndUpdate();
}

function KMSaveDocument(objHtmlDocument, objWizDocument) {
    if (!objWizDocument)
        return;
    if (!KMIsDocumentModified(objHtmlDocument))
        return;
    try {
        var progress = objApp.CreateWizObject("WizKMControls.WizProgressWindow");
        progress.Title = "Saving document";
        progress.Text = objWizDocument.Title;
        progress.Show();
        progress.Max = 1;
        //
        KMSaveDocumentCore(objHtmlDocument, objWizDocument);
        //
        progress.Hide();
        progress.Destroy();
    }
    catch (err) {
        WizAlert(err.message);
        //
    }
}

function KMOnDocumentBeforeEdit(objHtmlDocument, objWizDocument) {
    KMSaveDocument(objHtmlDocument, objWizDocument);
}
function KMOnDocumentBeforeClose(objHtmlDocument, objWizDocument) {
    KMSaveDocument(objHtmlDocument, objWizDocument);
}


function KMOnDocumentBeforeChange(objHtmlDocument, objWizDocumentOld, objWizDocumentNew) {
    KMSaveDocument(objHtmlDocument, objWizDocumentOld);
}

function UpdateButtonStatus() {
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var languageFileName = pluginPath + "plugin.ini";
    var strRead = objApp.LoadStringFromFile(languageFileName, "strRead");
    
    var objDoc = objApp.Window.CurrentDocument;
    
    var objDB = objApp.Database;
    var pNote_title = strRead + "(" + (objDoc.ReadCount-1) + ")";
    
    switch (objDoc.ParamValue("Rate")) {
    case "1":
        pNote_title = pNote_title + "|★☆☆☆☆";
        break;
    case "2":
        pNote_title = pNote_title + "|★★☆☆☆";
        break;
    case "3":
        pNote_title = pNote_title + "|★★★☆☆";
        break;
    case "4":
        pNote_title = pNote_title + "|★★★★☆";
        break;
    case "5":
        pNote_title = pNote_title + "|★★★★★";
        break;
    default:
        break;
    }
    objDB.Meta("TEST","MSG_500") = pNote_title;
    objWindow.RemoveToolButton("main", "KMHelperExButton");
    objWindow.AddToolButton("main", "KMHelperExButton", "【" + pNote_title + "】", "", "OnKMHelperExButtonClicked");
}

////////////////////////////////////////////////////////////////////////////////

function update_version() {
    if (objDB.Meta("keyword_HL","keyword_flag").value != null || 
        objDB.Meta("keyword_HL","keyword_flag").value != "" ) {
        objDB.Meta("wizhelp_parm","keyword_flag").value = objDB.Meta("keyword_HL","keyword_flag").value;
        //objDB.Meta("keyword_HL","keyword_flag").delete;
    }
}


function initParam() {
    if (objDB.Meta("wizhelp_parm","keyword_flag")=="" || objDB.Meta("wizhelp_parm","keyword_flag") == null ) {
        objDB.Meta("wizhelp_parm","keyword_flag") = "1";
    }
    if (objDB.Meta("wizhelp_parm","smarttag_flag")=="" || objDB.Meta("wizhelp_parm","smarttag_flag") == null ) {
        objDB.Meta("wizhelp_parm","smarttag_flag") = "1";
    }
}


function initEvents() {
/*
向Wiz注册一个事件，响应文档完成的消息。在Wiz内打开一个html文件的时候（例如阅读文档），如果Html文件打开完成，则调用这个方法。
*/

    eventsHtmlDocumentComplete.add(KMOnHtmlDocumentComplete);
    
    eventsDocumentBeforeEdit.add(KMOnDocumentBeforeEdit);
    eventsTabClose.add(KMOnDocumentBeforeClose);
    eventsClose.add(KMOnDocumentBeforeClose);
    eventsDocumentBeforeChange.add(KMOnDocumentBeforeChange);

}

function initPage() {
    initEvents();
}


function Main() {
    update_version();
    initParam();
    initPage();
}

////===========================================================================
//// 屏幕取词代码开始
//
// 生成取词窗口
function KMGetSearchWordWindow(doc, create) {
    var div = doc.getElementById("WizKMSearchWordDivID");
    if (div != null)
        return div;
    if (!create)
        return null;
    //
    div = doc.createElement("DIV");
    div.style.cssText = "padding:0; margin:0; position:absolute; z-index:100001; width:500px; height: 400px; display:none; background-color:window;border: solid 1px #6100C1;filter: progid:DXImageTransform.Microsoft.Shadow(color=#999999,direction=135,strength=5);";
    div.id = "WizKMSearchWordDivID";
    //
    return doc.body.appendChild(div);
}
//
// 关闭取词窗口
function KMCloseSearchWordWindow(doc, menu) {
    if (!menu) {
        menu = KMGetSearchWordWindow(doc, false);
    }
    if (!menu)
        return;
    if (menu.style.display == "none")
        return;
    menu.style.display = "none";
}
//
// 本地数据库及在线取词
function KMOnSearchWordClick() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var languageFileName = pluginPath + "plugin.ini";
    if (!doc)
        return;
    //
    var objSmartTag = KMGetSmartTagWindow(doc, false);
    if (!objSmartTag)
        return;
    //
    var divSearchWordWindow = KMGetSearchWordWindow(doc, true);
    if (!divSearchWordWindow)
        return;
    //
    KMCloseCommentWindow();
    //
    divSearchWordWindow.style.left = objSmartTag.offsetLeft + "px";
    divSearchWordWindow.style.top = (objSmartTag.offsetTop + objSmartTag.offsetHeight) + "px";
	divSearchWordWindow.innerHTML = "<table><tbody><tr><td id='tdKMSWWbtn0'></td><td id='tdKMSWWbtn1'></td><td id='tdKMSWWbtn2'></td></tr></tbody></table>";
    //
    var strWord = KMSmartTagGetSelectionText();
    if (!strWord || strWord == "") { return; }
    strWord = KMTextToSingleLine(strWord);
	//
    // 在线取词开始
	var htmlOnlineDicts = "";
	//htmlOnlineDicts += KMGetOnlineDictCn(strWord);
	htmlOnlineDicts += KMGetOnlineYoudao(strWord);
	if (htmlOnlineDicts != "") {
		var divKMDictOnline = doc.createElement("div");
		divKMDictOnline.id = "WizKMDictOnlineDivID";
		divKMDictOnline.style.margin = "3px";
		divKMDictOnline.innerHTML = htmlOnlineDicts;
		divKMDictOnline.style.display = "";
		//
		divSearchWordWindow.style.overflow = "scroll";
		divSearchWordWindow.appendChild(divKMDictOnline);
		divSearchWordWindow.style.display = "";
	}
    //
	// 本地词库取词开始
	var htmlLocalDict = "";
	htmlLocalDict += KMGetLocalDict(strWord);
	if (htmlLocalDict != "") {
		var divKMDictLocal = doc.createElement("div");
		divKMDictLocal.id = "WizKMDictLocalDivID";
		divKMDictLocal.style.display = "";
		divKMDictLocal.style.margin = "3px";
		divKMDictLocal.innerHTML = htmlLocalDict;
		//
		divSearchWordWindow.appendChild(divKMDictLocal);
		divSearchWordWindow.style.display = "";
	}
	//
	// 生成按钮
	if (htmlOnlineDicts!="" || htmlLocalDict!="") {
		// 生成保存精简取词结果为注释的按钮
		var btnSaveToComment = doc.createElement("span");
		btnSaveToComment.style.cssText = KMGetFlashMenuStyle1();
		btnSaveToComment.innerText = "[ "+objApp.LoadStringFromFile(languageFileName, "strKMSWasComment1")+" ]";
		btnSaveToComment.attachEvent("onclick", KMSaveWordToComment);
		btnSaveToComment.attachEvent("onmouseover", KMFlashMenuItemMouseOver);
		btnSaveToComment.attachEvent("onmouseout", KMFlashMenuItemMouseOut);
		doc.getElementById("tdKMSWWbtn1").appendChild(btnSaveToComment);
		// 生成保存所以取词结果为注释的按钮
		var btnSaveToComment2 = doc.createElement("span");
		btnSaveToComment2.style.cssText = KMGetFlashMenuStyle1();
		btnSaveToComment2.innerText = "[ "+objApp.LoadStringFromFile(languageFileName, "strKMSWasComment2")+" ]";
		btnSaveToComment2.attachEvent("onclick", KMSaveWordToComment2);
		btnSaveToComment2.attachEvent("onmouseover", KMFlashMenuItemMouseOver);
		btnSaveToComment2.attachEvent("onmouseout", KMFlashMenuItemMouseOut);
		doc.getElementById("tdKMSWWbtn2").appendChild(btnSaveToComment2);
	}
	//
	// 在线取得单词发声开始
	var urlAudio = KMGetOnlineDictCn(strWord);
	if (urlAudio != "") {
    	//var htmlAudio = "<embed src='" + urlAudio + "' width=\"50%\" height=\"50%\"></embed>"
		//var htmlAudio = "<audio src='" + urlAudio + "' controls autoplay>not supported</audio>";
		//var htmlAudio = "<img dynsrc='" + urlAudio + "' src='"+pluginPath+"audio.gif' alt='朗读单词' border='0' />";
		var htmlAudio = "<object id='KMDictAudioObjID' style='display:none' classid=\"clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95\">";
		htmlAudio += "<param name=\"AutoStart\" value=\"1\" />";
		htmlAudio += "<param name=\"FileName\" value=\"" + urlAudio + "\" />";
		htmlAudio += "</object>";
		//htmlAudio += "<img src='"+pluginPath+"audio.gif' alt='朗读单词' border='0' onclick='KMDictPlayAudio()' />";
		doc.getElementById("tdKMSWWbtn0").innerHTML = htmlAudio;
		//生成朗读单词的按钮
	    var btnSaveToComment0 = doc.createElement("span");
		btnSaveToComment0.style.cssText = KMGetFlashMenuStyle1();
		btnSaveToComment0.innerText = "[ "+objApp.LoadStringFromFile(languageFileName, "strKMSWasComment0")+" ]";
		btnSaveToComment0.attachEvent("onclick", KMDictPlayAudio);
		btnSaveToComment0.attachEvent("onmouseover", KMFlashMenuItemMouseOver);
		btnSaveToComment0.attachEvent("onmouseout", KMFlashMenuItemMouseOut);
		doc.getElementById("tdKMSWWbtn0").appendChild(btnSaveToComment0);
	}
	KMAutoCloseSearchWordWindow();
}
//
//
function KMDictPlayAudio() {
	var doc = objWindow.CurrentDocumentHtmlDocument;
	var objDictAudio = doc.getElementById("KMDictAudioObjID");
	objDictAudio.play();
}
//
//// 本地取词
function KMGetLocalDict(strWord) {
	var divHTML = "";
	var ConnDB = objApp.CreateActiveXObject("ADODB.Connection"); // 使用ADO的Connection对象打开数据库接口
	if (ConnDB) {
		var NameDB = "KMdictionary.mdb "; // Access数据库名
		var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
		var noteFileName = pluginPath + NameDB;
		var dbcon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " + noteFileName; // 操作指定数据库,Js使用相对地址
		ConnDB.Open(dbcon);
		var Rs = objApp.CreateActiveXObject("ADODB.Recordset");
		var strSQL = "select * from dictionary where word=\"" + strWord + "\"";
		Rs.Open(strSQL,ConnDB,1,3);
		while(!Rs.EOF) {
			divHTML += "<div style = 'font-size:9pt; font-style:italic; margin-top:15px;'> ========== <创世纪英语 本地词典> ========== </div>";
			divHTML += "<span id='DictLocal_pron' style='font-size:1em; font-weight:bold;'> [<label  style=\"font-family: 'Kingsoft Phonetic Plain'\">" + Rs("pron") + "</label>] </span>";
			divHTML += "<span style='font-size:1em; font-style:italic;'> syl: </span> ";
			divHTML += "<span id='DictLocal_ping' style='font-size:1em; font-weight:bold;'>" + Rs("ping") + "</span><br>";
			divHTML += "<span id='DictLocal_explain' style='font-size:1em;'>" + Rs("explain") + "</span><br>";
			divHTML += "<span id='DictLocal_difficult' style='font-size:0.8em;'>难度：【" + parseFloat(Rs("difficult")).toPrecision(2) + "】</span><br>";
			divHTML += "<span id='DictLocal_class' style='font-size:0.8em;'>类别：" + Rs("class") + "</span>";
			Rs.MoveNext;
		}
		ConnDB.Close(); //关闭数据库连接
	}
	return divHTML;
}
//
//// Youdao 在线取词开始
function KMGetOnlineYoudao(strWord) {
	var strHTML = "";
	strSearchWordURL = "http://fanyi.youdao.com/openapi.do?keyfrom=Wizhelper&key=342921866&type=data&doctype=xml&version=1.1&q=" + strWord;
	var xmlhttp_request = false;
	try { xmlhttp_request = objApp.CreateActiveXObject("Msxml2.XMLHTTP"); }
	catch(e) {
		try { xmlhttp_request = objApp.CreateActiveXObject("Microsoft.XMLHTTP"); }
		catch(e) { xmlhttp_request = false; }
	}
	if (xmlhttp_request) {
		xmlhttp_request.onreadystatechange = function(){
			if (xmlhttp_request.readyState == 4) {
				if (xmlhttp_request.status == 200) {
					var docXML = xmlhttp_request.responseXML.documentElement;
					var errorCode = docXML.getElementsByTagName("errorCode")[0].firstChild.data;
					if (errorCode == 0) {
						// Youdao Basic
						var objbasic = docXML.getElementsByTagName("basic")[0];
						if (objbasic) {
							strHTML += " <div style = 'font-size:9pt; font-style:italic; margin-top:15px;'> ========== <有道词典-基本词典> ========== </div>";
							strHTML += " <div id=\"WizKMDictYoudaoBasicDivID\"> ";
								var strPhonetic = objbasic.getElementsByTagName("phonetic")[0].firstChild.data;
								strHTML += " <b>[ " + strPhonetic + " ]</b> ";
								var objExplains = objbasic.getElementsByTagName("explains")[0].childNodes;
								for (var j=0; j<objExplains.length;j++) {
									var strExplainj = objExplains[j].firstChild.data;
									strHTML += " <br> " + strExplainj;
								}
							strHTML += " </div> ";
						}
						// Youdao Web
						var objweb = docXML.getElementsByTagName("web")[0];
						if (objweb) {
							strHTML += " <div style = 'font-size:9pt; font-style:italic; margin-top:15px;'> ========== <有道词典-网络释义> ========== </div>";
							strHTML += " <div id=\"WizKMDictYoudaoWebDivID\"> ";
								var objExplains = objweb.getElementsByTagName("explain");
								for (var j=0; j<objExplains.length;j++) {
									var strKeyj = objExplains[j].getElementsByTagName("key")[0].firstChild.data;
									strHTML += "<b>" + strKeyj + "</b> <br> ";
									var objValuesj = objExplains[j].getElementsByTagName("value")[0].childNodes;
									for (var k=0; k<objValuesj.length; k++) {
										var strValuek = objValuesj[k].firstChild.data;
										strHTML += strValuek + "; ";
									}
									strHTML += " <br> ";
								}
							strHTML += " </div> ";
						}
						// Youdao Translation
						var strTranslation = docXML.getElementsByTagName("paragraph")[0].firstChild.data;
						strHTML += " <div style = 'font-size:9pt; font-style:italic; margin-top:15px;'> ========== <有道在线翻译> ========== </div>";
						strHTML += " <div id=\"WizKMDictYoudaoTransDivID\"> ";
							strHTML += strTranslation;
						strHTML += " </div> ";
					}
				}
			}
		}
		xmlhttp_request.open("GET", strSearchWordURL, false);
		xmlhttp_request.setRequestHeader("Charset","GB2312");
		xmlhttp_request.setRequestHeader("Content-Type","text/xml");
		xmlhttp_request.send(null);
	}
	return strHTML;
}
//
//// Dict.cn 在线取词开始
function KMGetOnlineDictCn(strWord) {
	var strHTML = "";
	var strSearchWordURL = "http://dict.cn/ws.php?utf8=true&q=" + strWord;
	var xmlhttp_request = false;
	try { xmlhttp_request = objApp.CreateActiveXObject("Msxml2.XMLHTTP"); }
	catch(e) {
		try { xmlhttp_request = objApp.CreateActiveXObject("Microsoft.XMLHTTP"); }
		catch(e) { xmlhttp_request = false; }
	}
	if (xmlhttp_request) {
		xmlhttp_request.onreadystatechange = function(){
			if (xmlhttp_request.readyState == 4) {
				if (xmlhttp_request.status == 200) {
					var xmlText = xmlhttp_request.responseText;
					var iStart = xmlText.indexOf("<audio>")+7;
					var iStop = xmlText.indexOf("</audio>");
					var strURL = xmlText.substring(iStart,iStop);
					strHTML += strURL;
   /*                 var objDocXML = CreateXMLDocument(xmlhttp_request.responseText);*/
					//var objDict = objDocXML.getElementsByTagName("dict")[0];
					//var strKey = objDict.getElementsByTagName("key")[0].firstChild.data;
					////WizAlert(strKey);
					//var strPhonetic = objDict.getElementsByTagName("pron")[0].firstChild.data;
					//var urlAudio = objDict.getElementsByTagName("audio")[0].firstChild.data;
					//var strExplain = objDict.getElementsByTagName("def")[0].firstChild.data;
					//strHTML += " <h5>  --- Dict.cn 海词  --- </h5>";
					//strHTML += " <b>" + strKey + "</b> ";
					//strHTML += " <br><b>[ " + strPhonetic + " ]</b> ";
					//strHTML += " <a href='" + urlAudio + "'><img src='"+pluginPath+"audio.gif' alt=朗读单词 border='0' /></a> ";;
					//strHTML += " <br> " + strExplain;
					////WizAlert(strKey);
					//var objSentences = objDict.getElementsByTagName("sent");
					//if (objSentences){
						//strHTML += " <br><br> --- 例句 ---";
					//}
					//for (var j=0; j<objSentences.length;j++) {
						//var strSentenceEnj = objSentences[j].getElementsByTagName("orig")[0].firstChild.data;
						//var strSentenceCnj = objSentences[j].getElementsByTagName("trans")[0].firstChild.data;
						//strHTML += " <br> " + strSentenceEnj;
						//strHTML += " <br> " + strSentenceCnj + " <br> ";
					//}
					/*//WizAlert(strHTML);*/
				}
			}
		}
		xmlhttp_request.open("GET", strSearchWordURL, false);
		xmlhttp_request.setRequestHeader("Charset","GB2312");
		xmlhttp_request.setRequestHeader("Content-Type","text/xml");
		xmlhttp_request.send(null);
	}
	return strHTML;
}
//
// 保存精简取词结果为注释
function KMSaveWordToComment() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc) { return; }
    if (!g_KMSelection) { return; }
    var html = g_KMSelection.htmlText;
    //
    var id = "WizKMComment_" + KMGetRandomInt();
    //
    var newHtml = "<span id=\"" + id + "\" style=\"border-bottom: 2px #ff0000 dashed\" >" + html + "</span>";
    //
    var objDictLocalDiv = doc.getElementById("WizKMDictLocalDivID");
	var objDictYoudaoDiv = doc.getElementById("WizKMDictYoudaoBasicDivID");
	//
	var objDictAudio = doc.getElementById("KMDictAudioObjID");
	var urlAudio = "";
	urlAudio += objDictAudio.lastChild.value;

    var strWordToComment = "";
	if (objDictYoudaoDiv) {
    	strWordToComment += objDictYoudaoDiv.innerHTML;
	}
	else if (objDictLocalDiv) {
    	for (var j=2; j<6;j++) {
    		strWordToComment += objDictLocalDiv.childNodes[j].outerHTML;
    	}
	}
    try {
        g_KMSelection.pasteHTML(newHtml);
        var objInlineComment = doc.getElementById(id);
        if (objInlineComment) {
            objInlineComment.title = strWordToComment;
			if (urlAudio != "") {
				objInlineComment.audioURL = urlAudio;
			}
            KMCommentAttachEvents(objInlineComment);
        }
    }
    catch (err) {
    }
	KMCloseSearchWordWindow(doc, null);
    KMSetDocumentModified(doc);
}
// 保存全部取词结果为注释
function KMSaveWordToComment2() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc) { return; }
    if (!g_KMSelection) { return; }
    var html = g_KMSelection.htmlText;
    //
    var id = "WizKMComment_" + KMGetRandomInt();
    //
    var newHtml = "<span id=\"" + id + "\" style=\"border-bottom: 2px #ff0000 dashed\" >" + html + "</span>";
    //
    var objDictLocalDiv = doc.getElementById("WizKMDictLocalDivID");
    var objDictOnlineDiv = doc.getElementById("WizKMDictOnlineDivID");
    try {
        g_KMSelection.pasteHTML(newHtml);
        var objInlineComment = doc.getElementById(id);
        if (objInlineComment) {
            objInlineComment.title =  objDictOnlineDiv.innerHTML + objDictLocalDiv.innerHTML;
            KMCommentAttachEvents(objInlineComment);
        }
    }
    catch (err) {
    }
	KMCloseSearchWordWindow(doc, null);
    KMSetDocumentModified(doc);
}
//
//
function KMIsSearchWord(elem) {
    if (!elem.id)
        return false;
    if (elem.id.indexOf("WizKMComment_") != 0)
        return false;
    return true;
}
function KMGetParentSearchWord(elem) {
    while (elem != null) {
        if (KMIsSearchWord(elem))
            return elem;
        elem = elem.parentElement;
    }
    return null;
}
//
//当鼠标移开时自动隐藏标题窗口
function KMAutoCloseSearchWordWindowTimer() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    //
    var objSearchWordWindow = KMGetSearchWordWindow(doc, false);
    if (!objSearchWordWindow)
        return;
    //
    if (objSearchWordWindow.style.display != "none") {
        var arr = [doc.activeElement, doc.elementFromPoint(g_KMMousePos.x, g_KMMousePos.y)];
        for (var i = 0; i < arr.length; i++) {
            var elem = arr[i];
            if (!elem)
                continue;
            if (null != KMGetParentSearchWord(elem))
                return;
            if (KMIsChildNodeOf(elem, "WizKMSearchWordDivID"))
                return;
        }
        objSearchWordWindow.style.display = "none";
    }
    //
    objWindow.RemoveTimer("KMAutoCloseSearchWordWindowTimer");
}
//
//
function KMAutoCloseSearchWordWindow() {
    objWindow.RemoveTimer("KMAutoCloseSearchWordWindowTimer");
    objWindow.AddTimer("KMAutoCloseSearchWordWindowTimer", 1000);
}
//
//// 屏幕取词代码结束
////===========================================================================
////
////===========================================================================
////添加链接到【文件夹|标签|样式|文件】代码开始
//点击【链接到...】按钮
function KMOnLinkToClick() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var languageFileName = pluginPath + "plugin.ini";
    if (!doc)
        return;
    //
    var objSmartTag = KMGetSmartTagWindow(doc, false);
    if (!objSmartTag)
        return;
    //
    var divLinkToFolderWindow = KMGetLinkToWindow(doc, true);
    if (!divLinkToFolderWindow)
        return;
    //
    KMCloseCommentWindow();
    //
    divLinkToFolderWindow.style.left = objSmartTag.offsetLeft + "px";
    divLinkToFolderWindow.style.top = (objSmartTag.offsetTop + objSmartTag.offsetHeight) + "px";
    divLinkToFolderWindow.innerHTML = "";
    divLinkToFolderWindow.style.display = "";
    //
    var strWord = KMSmartTagGetSelectionText();
    if (!strWord || strWord == "") { return; }
    //
    if (objWindow.CategoryCtrl.SelectedFolder){
        var btnAddFolder = doc.createElement("span");
        btnAddFolder.style.cssText = KMGetFlashMenuStyle1();
        btnAddFolder.innerText = objApp.LoadStringFromFile(languageFileName, "strLinkToFolder");
        btnAddFolder.attachEvent("onclick", KMLinkToFolder);
        btnAddFolder.attachEvent("onmouseover", KMFlashMenuItemMouseOver);
        btnAddFolder.attachEvent("onmouseout", KMFlashMenuItemMouseOut);
        divLinkToFolderWindow.appendChild(btnAddFolder);
    }
    if (objWindow.CategoryCtrl.SelectedTags.Count>0){
        var btnAddFolder = doc.createElement("span");
        btnAddFolder.style.cssText = KMGetFlashMenuStyle1();
        btnAddFolder.innerText = objApp.LoadStringFromFile(languageFileName, "strLinkToTags");
        btnAddFolder.attachEvent("onclick", KMLinkToTags);
        btnAddFolder.attachEvent("onmouseover", KMFlashMenuItemMouseOver);
        btnAddFolder.attachEvent("onmouseout", KMFlashMenuItemMouseOut);
        divLinkToFolderWindow.appendChild(btnAddFolder);
    }
    if (objWindow.CategoryCtrl.SelectedStyle){
        var btnAddFolder = doc.createElement("span");
        btnAddFolder.style.cssText = KMGetFlashMenuStyle1();
        btnAddFolder.innerText = objApp.LoadStringFromFile(languageFileName, "strLinkToStyle");
        btnAddFolder.attachEvent("onclick", KMLinkToStyle);
        btnAddFolder.attachEvent("onmouseover", KMFlashMenuItemMouseOver);
        btnAddFolder.attachEvent("onmouseout", KMFlashMenuItemMouseOut);
        divLinkToFolderWindow.appendChild(btnAddFolder);
    }
    if (objWindow.DocumentsCtrl.SelectedDocuments.Count>0){
        var btnAddFolder = doc.createElement("span");
        btnAddFolder.style.cssText = KMGetFlashMenuStyle1();
        btnAddFolder.innerText = objApp.LoadStringFromFile(languageFileName, "strLinkToDocs");
        btnAddFolder.attachEvent("onclick", KMLinkToDocs);
        btnAddFolder.attachEvent("onmouseover", KMFlashMenuItemMouseOver);
        btnAddFolder.attachEvent("onmouseout", KMFlashMenuItemMouseOut);
        divLinkToFolderWindow.appendChild(btnAddFolder);
    }
    divLinkToFolderWindow.style.display = "";
    //
}
//
// 生成【链接到...】窗口
function KMGetLinkToWindow(doc, create) {
    var div = doc.getElementById("WizKMLinkToFolderDivID");
    if (div != null)
        return div;
    if (!create)
        return null;
    //
    div = doc.createElement("DIV");
    div.style.cssText = "padding:0; margin:0; position:absolute; z-index:100001; width:200px; background-color:window;border: solid 1px #6100C1;filter: progid:DXImageTransform.Microsoft.Shadow(color=#999999,direction=135,strength=5);";
    div.id = "WizKMLinkToFolderDivID";
    //
    return doc.body.appendChild(div);
}
//
// 关闭【链接到...】窗口
function KMCloseLinkToWindow(doc, menu) {
    if (!menu) {
        menu = KMGetLinkToWindow(doc, false);
    }
    if (!menu)
        return;
    if (menu.style.display == "none")
        return;
    menu.style.display = "none";
}
//
//【链接到选中文件夹】
function KMLinkToFolder() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    KMCloseLinkToWindow(doc, null);
    KMSetLinkToSelection("folder");
}
//【链接到选中标签】
function KMLinkToTags() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    KMCloseLinkToWindow(doc, null);
    KMSetLinkToSelection("tags");
}
//【链接到选中样式】
function KMLinkToStyle() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    KMCloseLinkToWindow(doc, null);
    KMSetLinkToSelection("style");
}
//【链接到选中的多个文件】
function KMLinkToDocs() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    KMCloseLinkToWindow(doc, null);
    KMSetLinkToSelection("docs");
}
//
//根据当前是否有【文件夹/标签/样式/文件】选中而生成相应按钮
function KMSetLinkToSelection(str) {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc) { return; }
    if (!g_KMSelection) { return; }
    var html = g_KMSelection.htmlText;
    //
    var id = "WizKMLink_" + KMGetRandomInt();
    //
    switch (str) {
        case "folder":
            var strFolderLocation = objWindow.CategoryCtrl.SelectedFolder.Location;
            var newHtml = "<span id=\"" + id + "\"; style=\"text-decoration:underline; color:blue;\" folderlocation=\"" + strFolderLocation + "\"; name=\"\">" + html + "</span>";
            break;
        case "tags":
            var objSelectedTags = objWindow.CategoryCtrl.SelectedTags;
            var strTagsGUID = "";
            for (var i=0; i<objSelectedTags.Count; i++){
                strTagsGUID += objSelectedTags.Item(i).GUID + ";";
            }
            var newHtml = "<span id=\"" + id + "\"; style=\"text-decoration:underline; color:blue;\" tagsguid=\"" + strTagsGUID + "\"; name=\"\">" + html + "</span>";
            break;
        case "style":
            var strStyleGUID = objWindow.CategoryCtrl.SelectedStyle.GUID;
            var newHtml = "<span id=\"" + id + "\"; style=\"text-decoration:underline; color:blue;\" styleguid=\"" + strStyleGUID + "\"; name=\"\">" + html + "</span>";
            break;
        case "docs":
            var objSelectedDocs = objWindow.DocumentsCtrl.SelectedDocuments;
            var strDocsGUID = "";
            for (var i=0; i<objSelectedDocs.Count; i++){
                strDocsGUID += objSelectedDocs.Item(i).GUID + ";";
            }
            var newHtml = "<span id=\"" + id + "\"; style=\"text-decoration:underline; color:blue;\" docsguid=\"" + strDocsGUID + "\"; name=\"\">" + html + "</span>";
            break;
        default:
        {}
    }
    //
    try {
        g_KMSelection.pasteHTML(newHtml);
        var objInlineLink = doc.getElementById(id);
        if (objInlineLink) {
            KMLinkToAttachEvents(objInlineLink);
        }
    }
    catch (err) {
    }
    KMSetDocumentModified(doc);
}
//添加鼠标事件
function KMAttachLinkToEvents(doc) {
    var arr = doc.getElementsByTagName("SPAN");
    for (var i = 0; i < arr.length; i++) {
        var elem = arr[i];
        if (!KMIsInlineLink(elem))
            continue;
        //
        KMLinkToAttachEvents(elem);
    }
}
function KMLinkToAttachEvents(obj) {
    obj.attachEvent("onclick", KMOnLinkToTextClicked);
    obj.attachEvent("onmouseover", KMOnLinkToTextMouseOver);
}
function KMIsInlineLink(elem) {
    if (!elem.id)
        return false;
    if (elem.id.indexOf("WizKMLink_") != 0)
        return false;
    return true;
}
function KMGetParentInlineLink(elem) {
    while (elem != null) {
        if (KMIsInlineLink(elem))
            return elem;
        //
        elem = elem.parentElement;
    }
    return null;
}
//点击链接文字筛选文档
function KMOnLinkToTextClicked() {
	var objDatabase = objApp.Database;
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    //
    var win = doc.parentWindow;
    if (!win)
        return;
    //
    var e = win.event;
    if (!e)
        return;
    //
    var elem = e.srcElement;
    if (!elem)
        return;
    //
    var objInlineLink = KMGetParentInlineLink(elem);
    if (!objInlineLink)
        return;
    //
    if (objInlineLink.folderlocation){
        var folderlocation = objInlineLink.folderlocation;
        var objFolder = objDatabase.GetFolderByLocation(folderlocation,true);
        if (!objFolder)
            return;
        objWindow.CategoryCtrl.SelectedFolder=objFolder;
    }
    if (objInlineLink.tagsguid){
		var objTags = objApp.CreateWizObject("WizKMCore.WizTagCollection");
        var listTagsGUID = objInlineLink.tagsguid.split(";");
        for (var i=0; i<listTagsGUID.length-1; i++){
            var objTag = objDatabase.TagFromGUID(listTagsGUID[i]);
            objTags.Add(objTag);
        }
        var objDocs = objDatabase.DocumentsFromTags(objTags);
        objWindow.DocumentsCtrl.SetDocuments(objDocs);
    }
    if (objInlineLink.styleguid){
        var styleguid = objInlineLink.styleguid;
        var objStyle = objDatabase.StyleFromGUID(styleguid);
        if (!objStyle)
            return;
        objWindow.CategoryCtrl.SelectedStyle=objStyle;
    }
    if (objInlineLink.docsguid){
		var objDocs = objApp.CreateWizObject("WizKMCore.WizDocumentCollection");
        var listDocsGUID = objInlineLink.docsguid.split(";");
        for (var i=0; i<listDocsGUID.length-1; i++){
            var objDoc = objDatabase.DocumentFromGUID(listDocsGUID[i]);
            objDocs.Add(objDoc);
        }
        objWindow.DocumentsCtrl.SetDocuments(objDocs);
    }
}
///-----------------------------------------------------------------------
//删除链接部分代码开始
//鼠标浮停时显示【删除链接】窗口
function KMOnLinkToTextMouseOver() {
	var objDatabase = objApp.Database;
    var doc = objWindow.CurrentDocumentHtmlDocument;
    var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var languageFileName = pluginPath + "plugin.ini";
    if (!doc)
        return;
    //
    var win = doc.parentWindow;
    if (!win)
        return;
    //
    var e = win.event;
    if (!e)
        return;
    //
    var elem = e.srcElement;
    if (!elem)
        return;
    //
    var objInlineLink = KMGetParentInlineLink(elem);
    if (!objInlineLink)
        return;
    //
    var objEditLinkWindow = KMGetEditLinkWindow(doc, true);
    if (!objEditLinkWindow)
        return;
    var scrollX = Math.max(doc.body.scrollLeft, doc.documentElement.scrollLeft);
    var scrollY = Math.max(doc.body.scrollTop, doc.documentElement.scrollTop);
    var x = e.clientX + scrollX;
    var y = e.clientY + scrollY;
    //
    var widthEditLinkWindow = parseInt(objEditLinkWindow.style.width);
    //
    if (x + widthEditLinkWindow > doc.body.clientWidth) {
        x = doc.body.clientWidth - widthEditLinkWindow;
    }
    if (x < 0) {
        x = 0;
    }
    //
    y += 20;
    //
    objEditLinkWindow.style.left = x + "px";
    objEditLinkWindow.style.top = y + "px";
    objEditLinkWindow.innerHTML = "";
    objEditLinkWindow.setAttribute("KMLinkCurrentID", objInlineLink.id);
    //
    var btnDelLink = doc.createElement("span");
    btnDelLink.style.cssText = KMGetFlashMenuStyle1();
    btnDelLink.innerText = objApp.LoadStringFromFile(languageFileName, "strDelLink");
    btnDelLink.attachEvent("onclick", KMDelLink);
    btnDelLink.attachEvent("onmouseover", KMFlashMenuItemMouseOver);
    btnDelLink.attachEvent("onmouseout", KMFlashMenuItemMouseOut);
    objEditLinkWindow.appendChild(btnDelLink);
    objEditLinkWindow.style.display = "";
    KMAutoCloseEditLinkWindow();
}
//
// 生成【删除链接】窗口
function KMGetEditLinkWindow(doc, create) {
    var div = doc.getElementById("WizKMEditLinkToDivID");
    if (div != null)
        return div;
    if (!create)
        return null;
    //
    div = doc.createElement("DIV");
    div.style.cssText = "padding:0; margin:0; position:absolute; z-index:100001; width:168px; height: 30px; display:none; background-color:window;border: solid 1px #6100C1;filter: progid:DXImageTransform.Microsoft.Shadow(color=#999999,direction=135,strength=5);";
    div.id = "WizKMEditLinkToDivID";
    //
    return doc.body.appendChild(div);
}
//
// 关闭【删除链接】窗口
function KMCloseEditLinkToWindow(doc, menu) {
    if (!menu) {
        menu = KMGetEditLinkWindow(doc, false);
    }
    if (!menu)
        return;
    if (menu.style.display == "none")
        return;
    menu.style.display = "none";
}
//
//删除链接
function KMDelLink() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    var objEditLinkWindow = KMGetEditLinkWindow(doc, false);
    if (!objEditLinkWindow)
        return;
    //
    var objInlineLink = KMGetCurrentInlineLink(doc, objEditLinkWindow);
    //
    if (objInlineLink) {
        objInlineLink.removeNode(false);
        KMSetDocumentModified(doc);
    }
    KMCloseEditLinkToWindow(doc, null);
}
//获取当前链接id
function KMGetCurrentInlineLink(doc, objEditLinkWindow) {
    if (!objEditLinkWindow) {
        objEditLinkWindow = KMGetEditLinkWindow(doc, false);
    }
    if (!objEditLinkWindow)
        return null;
    var inlineLinkID = objEditLinkWindow.getAttribute("KMLinkCurrentID");
    if (inlineLinkID == null || inlineLinkID == "")
        return null;
    //
    var objInlineLink = doc.getElementById(inlineLinkID);
    if (!objInlineLink)
        return null;
    return objInlineLink;
}
//当鼠标移开时自动隐藏取消链接窗口
function KMAutoCloseEditLinkWindowTimer() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    //
    var objEditLinkWindow = KMGetEditLinkWindow(doc, false);
    if (!objEditLinkWindow)
        return;
    //
    if (objEditLinkWindow.style.display != "none") {
        var arr = [doc.activeElement, doc.elementFromPoint(g_KMMousePos.x, g_KMMousePos.y)];
        for (var i = 0; i < arr.length; i++) {
            var elem = arr[i];
            if (!elem)
                continue;
            if (null != KMGetParentInlineLink(elem))
                return;
            if (KMIsChildNodeOf(elem, "WizKMEditLinkToDivID"))
                return;
        }
        objEditLinkWindow.style.display = "none";
    }
    //
    objWindow.RemoveTimer("KMAutoCloseEditLinkWindowTimer");
}
//
function KMAutoCloseEditLinkWindow() {
    objWindow.RemoveTimer("KMAutoCloseEditLinkWindowTimer");
    objWindow.AddTimer("KMAutoCloseEditLinkWindowTimer", 1000);
}
//删除链接部分代码结束
///-----------------------------------------------------------------------
////添加链接到【文件夹|标签|样式|文件】代码结束
////=================================================================================
////
////
////=================================================================================
////设置为目录项代码开始
//
function KMOnSetAsContentClick() {
	var objDatabase = objApp.Database;
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc) { return; }
    var objSmartTag = KMGetSmartTagWindow(doc, false);
    if (!objSmartTag) { return; }
    var divContentWindow = KMGetContentWindow(doc, true);
    if (!divContentWindow) { return; }
    //
    KMCloseFlashMenu(doc, null);
    KMCloseSearchWordWindow(doc, null);
    KMCloseCommentWindow();
    KMCloseLinkToWindow(doc, null);
    KMCloseEditLinkToWindow(doc, null);
    //
    divContentWindow.style.left = objSmartTag.offsetLeft + "px";
    divContentWindow.style.top = (objSmartTag.offsetTop + objSmartTag.offsetHeight) + "px";
    divContentWindow.style.display = "";
    //
    var strContentTitle = KMSmartTagGetSelectionText();
    if (!strContentTitle || strContentTitle == "")
        return;
    var metaName = "varKMContents";
    var metasContents = objDatabase.MetasByName(metaName);
    var chkBackToTop = doc.getElementById("KMContentBackToTopID");
	if (metasContents && objDatabase.Meta(metaName,"isAddBackToTop")=="1") {
	    chkBackToTop.setAttribute("checked", true);
    }
    else {
        chkBackToTop.setAttribute("checked", false);
        objDatabase.Meta(metaName,"isAddBackToTop") = "0";
    }
    //
    KMAutoCloseContentWindow();
}
//
// 生成【设置为目录项...】窗口
function KMGetContentWindow(doc, create) {
	var pluginPath = objApp.GetPluginPathByScriptFileName("km.js");
    var languageFileName = pluginPath + "plugin.ini";
    var div = doc.getElementById("WizKMContentDivID");
    if (div != null)
        return div;
    if (!create)
        return null;
    //
    div = doc.createElement("DIV");
    div.style.cssText = "padding:0; margin:0; position:absolute; z-index:100001; width:200px; display:none; background-color:window;border: solid 1px #6100C1;filter: progid:DXImageTransform.Microsoft.Shadow(color=#999999,direction=135,strength=5);";
    div.id = "WizKMContentDivID";
    //
    KMAddFlashMenuItem2(doc, div, "strSelAsContent1", KMFlashSelAsContent1);
    KMAddFlashMenuItem2(doc, div, "strSelAsContent2", KMFlashSelAsContent2);
    KMAddFlashMenuItem2(doc, div, "strSelAsContent3", KMFlashSelAsContent3);
    KMAddFlashMenuItem2(doc, div, "strSelAsContent4", KMFlashSelAsContent4);
    KMAddFlashMenuItem2(doc, div, "strSelAsContent5", KMFlashSelAsContent5);
    KMAddFlashMenuItem2(doc, div, "strSelUnContent", KMFlashSelUnContent);
    KMAddFlashMenuItem2(doc, div, "strBackToTop", KMAddBackToTopLink);
    //
	var chkBackToTop = doc.createElement("input");
	chkBackToTop.type = "checkbox";
	chkBackToTop.id="KMContentBackToTopID";
	chkBackToTop.attachEvent("onclick", KMSetMetaBackToTop);
	div.appendChild(chkBackToTop);
	var labelBackToTop = doc.createElement("label");
	labelBackToTop.innerText = objApp.LoadStringFromFile(languageFileName, "strBackToTop");
	labelBackToTop.style.cssText = "padding:2; text-align: left; text-decoration: none; font-family:arial; font-size:9pt; color: #000000;";
	div.appendChild(labelBackToTop);
    //
    return doc.body.appendChild(div);
}
//
// 保存是否添加【链接到页面顶部】
function KMSetMetaBackToTop(){
	var objDatabase = objApp.Database;
	var doc = objWindow.CurrentDocumentHtmlDocument;
    var chkBackToTop = doc.getElementById("KMContentBackToTopID");
    if (chkBackToTop.checked) { objDatabase.Meta("varKMContents","isAddBackToTop") = "1"; }
    else { objDatabase.Meta("varKMContents","isAddBackToTop") = "0"; }
}
//
// 关闭【设置为目录项...】窗口
function KMCloseContentWindow() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc) { return; }
    objContentWindow = KMGetContentWindow(doc, false);
    if (!objContentWindow) { return; }
    objContentWindow.style.display = "none";
}
//
// 保存目录项
function KMFlashSelAsContent1() {
	KMFlashSelAsContent(1);
}
function KMFlashSelAsContent2() {
	KMFlashSelAsContent(2);
}
function KMFlashSelAsContent3() {
	KMFlashSelAsContent(3);
}
function KMFlashSelAsContent4() {
	KMFlashSelAsContent(4);
}
function KMFlashSelAsContent5() {
	KMFlashSelAsContent(5);
}
function KMFlashSelAsContent(intContentClass) {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc) { return; }
    if (!g_KMSelection) { return; }
    //
    var objContent = KMGetParentContent(g_KMSelection.parentElement());
    if (objContent) {
     	objContent.style.fontSize = (20-2*intContentClass).toString() + "pt";
     	objContent.KMContentClass = intContentClass;
    }
    else {
    	var  html = g_KMSelection.htmlText;
	    var id = "WizKMContent_" + KMGetRandomInt();
		var name = KMContentGetName(g_KMSelection.text);
    	var objdiv = doc.createElement("div");
	    objdiv.innerHTML = html;
	    var chkBackToTop = doc.getElementById("KMContentBackToTopID");
	    if (chkBackToTop && chkBackToTop.checked) {
	    	var objAPageTop = doc.getElementById("KMContentPageTopID");
	    	if (objAPageTop) {
	    		if (doc.body.firstChild.id!="KMContentPageTopID"){
	    			objAPageTop.removeNode(true);
	    			objAPageTop = doc.createElement("a");
					objAPageTop.id="KMContentPageTopID";
					objAPageTop.name = "KMContentPageTopID";
				    doc.body.insertBefore(objAPageTop,doc.body.firstChild);
	    		}
	    	}
	    	else {
    			objAPageTop = doc.createElement("a");
				objAPageTop.id="KMContentPageTopID";
				objAPageTop.name = "KMContentPageTopID";
			    doc.body.insertBefore(objAPageTop,doc.body.firstChild);
	    	}
	    	var objABackToTop = doc.createElement("a");
	    	objABackToTop.id =  id + "_Top";
	    	objABackToTop.style.cssText = "cursor:hand; color:grey;";
	    	objABackToTop.href = "#KMContentPageTopID";
	    	objABackToTop.innerText = "[^]";
	    	if ("PDIVH1H2H3H4H5H6".indexOf(objdiv.lastChild.tagName) != -1) { objdiv.lastChild.appendChild(objABackToTop); }
	    	else if ("BR".indexOf(objdiv.lastChild.tagName) != -1) { objdiv.insertBefore(objABackToTop,objdiv.lastChild); }
	    	else { objdiv.appendChild(objABackToTop); }
		}
		html = objdiv.innerHTML;
		var newHtml = "<a id=\"" + id + "\"; style=\"font-weight:bold; font-size:" + (20-2*intContentClass) + "pt;\"  name=" + name + " KMContentClass=" + intContentClass + ">" + html + "</a>";
	    try { g_KMSelection.pasteHTML(newHtml); }
	    catch (err) { }
	}
    KMSetDocumentModified(doc);
    KMCloseContentWindow();
}
//
//取消目录项
function KMFlashSelUnContent() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc) { return; }
    if (!g_KMSelection) { return; }
    //
    var objContent = KMGetParentContent(g_KMSelection.parentElement());
    if (objContent) {
        objContent.removeNode(false);
        var ojbLinkTop = doc.getElementById(objContent.id + "_Top");
        if (ojbLinkTop) {
        	ojbLinkTop.removeNode(true);
        }
        KMSetDocumentModified(doc);
    }
    KMCloseContentWindow();
}
//
//
function KMIsContent(elem) {
    if (!elem.id)
        return false;
    if (elem.id.indexOf("WizKMContent_") != 0)
        return false;
    return true;
}
function KMGetParentContent(elem) {
    while (elem != null) {
        if (KMIsContent(elem))
            return elem;
        elem = elem.parentElement;
    }
    return null;
}
//
//当鼠标移开时自动隐藏标题窗口
function KMAutoCloseContentWindowTimer() {
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc)
        return;
    //
    var objContentWindow = KMGetContentWindow(doc, false);
    if (!objContentWindow)
        return;
    //
    if (objContentWindow.style.display != "none") {
        var arr = [doc.activeElement, doc.elementFromPoint(g_KMMousePos.x, g_KMMousePos.y)];
        for (var i = 0; i < arr.length; i++) {
            var elem = arr[i];
            if (!elem)
                continue;
            if (null != KMGetParentContent(elem))
                return;
            if (KMIsChildNodeOf(elem, "WizKMContentDivID"))
                return;
        }
        objContentWindow.style.display = "none";
    }
    //
    objWindow.RemoveTimer("KMAutoCloseContentWindowTimer");
}
//
//
function KMAutoCloseContentWindow() {
    objWindow.RemoveTimer("KMAutoCloseContentWindowTimer");
    objWindow.AddTimer("KMAutoCloseContentWindowTimer", 1000);
}
//
// 插入【回到页面顶端】漂浮连接
function KMAddBackToTopLink(){
    var doc = objWindow.CurrentDocumentHtmlDocument;
    if (!doc) { return; }
	var objDivBody = doc.getElementById('KMContentBodyDivID');
	if (objDivBody) { doc.body.innerHTML = objDivBody.innerHTML; }
	doc.body.style.cssText = "height:100%; overflow:auto;";
	//
	var objHTML = doc.getElementsByTagName('html')[0];
	objHTML.style.verflow = "hidden";
	//
	var objdivBackToTop = doc.createElement("div");
	objdivBackToTop.style.cssText = "position:absolute; right:5px; top:expression(eval(document.body.scrollTop + document.body.clientHeight - 50));";
	objdivBackToTop.id = "KMContentBackToTopDivID";
	objdivBackToTop.innerHTML = "<a href=\"#KMContentPageTopID\";>^</a>";
	doc.body.insertBefore(objdivBackToTop,doc.body.firstChild);
	//
	var objAPageTop = doc.getElementById("KMContentPageTopID");
	if (objAPageTop) { objAPageTop.removeNode(true); }
	objAPageTop = doc.createElement("a");
	objAPageTop.id="KMContentPageTopID";
	doc.body.insertBefore(objAPageTop,doc.body.firstChild);
	//
	KMSetDocumentModified(doc);
	KMCloseContentWindow();
}
//
//去除 Bookmark name 中的非法字符
function KMContentGetName(str) {
	return str.replace(/\"|'|`/g,"");
}
////
////设置为目录项代码结束
////=================================================================================
////////////////////////////////////////////////////////////////////////////////
var objApp = WizExplorerApp;
var objDB = objApp.Database;
Main();

