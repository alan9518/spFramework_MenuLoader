var jQuery = "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.0.2.min.js";
var sharingStatus = "";
var policy = "";
var statusBarBackground;
var hostUrl;
var subSiteUrl;
var currentUrl;
var loadPopup = "";
// Register script for MDS if possible
// Is MDS enabled?
if ("undefined" != typeof g_MinimalDownload && g_MinimalDownload && (window.location.pathname.toLowerCase()).endsWith("/_layouts/15/start.aspx") && "undefined" != typeof asyncDeltaManager) {
    // Register script for MDS if possible
    RegisterModuleInit("siteprivacy.js", JavaScript_Embed); //MDS registration
    JavaScript_Embed(); //non MDS run
} else {
    JavaScript_Embed();
}

function JavaScript_Embed() {

    loadScript(jQuery, function () {
        $(document).ready(function () {
            var message = "";

            
           
            // Execute status setter only after SP.JS has been loaded
            SP.SOD.executeOrDelayUntilScriptLoaded(function () {
                //alert("Site Collection: " + window.location.host + _spPageContextInfo.siteServerRelativeUrl);
                //alert("Sub Site: " + _spPageContextInfo.webAbsoluteUrl);

                hostUrl = window.location.host + _spPageContextInfo.siteServerRelativeUrl;
                subSiteUrl = _spPageContextInfo.webAbsoluteUrl;
                if (hostUrl != subSiteUrl) {
                    currentUrl = subSiteUrl;
                }
                else {
                    currentUrl = hostUrl;
                }

                var appweb = _spPageContextInfo.isAppWeb;
                if (appweb == false && appweb != 'undefined')
                {
                    getClassifier();
                }

            }, 'sp.js');
        });
    });
}

function getClassifier() {
    var currentUser;
    var clientContext = SP.ClientContext.get_current();
    var web = clientContext.get_web();
    currentUser = web.get_currentUser();
    var props = web.get_allProperties();
    clientContext.load(web);
    clientContext.load(currentUser);
    clientContext.load(web, 'EffectiveBasePermissions');
    clientContext.load(props);

    clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
        try{
            //policy = props.get_item('PolicyName');
            var isOwner = "No";
            if (web.get_effectiveBasePermissions().has(SP.PermissionKind.manageWeb)) {
                //User Has Edit Permissions
                isOwner = "Yes";
            }
            policy = "";
            if (policy.length != 0) {
                var sharingStatus1 = props.get_item("_site_props_MainOwner");
                var sharingStatus2 = props.get_item("_site_props_Type");
                var sharingStatus3 = props.get_item("_site_props_CreationDate");
                var sharingStatus4 = props.get_item("_site_props_BlockDate");
                var sharingStatus5 = props.get_item("_site_props_Compliance");

                getSiteSharingStatus();
            }
            else
            {
                policy = "No Site Policy";

                var sharingStatus1 = props.get_item("_site_props_MainOwner");
                var sharingStatus2 = props.get_item("_site_props_Type");
                var sharingStatus3 = props.get_item("_site_props_CreationDate");
                var sharingStatus4 = props.get_item("_site_props_BlockDate");
                var sharingStatus5 = props.get_item("_site_props_Compliance");
                var sharingStatus6 = props.get_item("_site_props_closed");
                var sharingStatus7 = props.get_item("_site_props_expirydate");
                var sharingStatus8 = props.get_item("_site_props_policy");
                var sharingStatus9 = props.get_item("_site_props_dbid");
                var currentUser = _spPageContextInfo.userLoginName;
                getSiteSharingStatus();

                setUI(sharingStatus1, sharingStatus2, sharingStatus3, sharingStatus4, sharingStatus5, sharingStatus6, sharingStatus7, sharingStatus8, sharingStatus9, currentUser, isOwner);
            }
        }
        catch(e){
            policy = "No Site Policy";
			
		    var sharingStatus1 = props.get_item("_site_props_MainOwner");
            var sharingStatus2 = props.get_item("_site_props_Type");
            var sharingStatus3 = props.get_item("_site_props_CreationDate");
            var sharingStatus4 = props.get_item("_site_props_BlockDate");
            var sharingStatus5 = props.get_item("_site_props_Compliance");
            var sharingStatus6 = props.get_item("_site_props_closed");
            var sharingStatus7 = props.get_item("_site_props_expirydate");
            var sharingStatus8 = props.get_item("_site_props_policy");
            var sharingStatus9 = props.get_item("_site_props_dbid");
            var currentUser = _spPageContextInfo.userLoginName;
            getSiteSharingStatus();
            
            setUI(sharingStatus1, sharingStatus2, sharingStatus3, sharingStatus4, sharingStatus5, sharingStatus6, sharingStatus7, sharingStatus8, sharingStatus9, currentUser, isOwner);
        }
        
    }));

    //return policy;
}

function SetStatusBar(message, bgColor) {
    strUpdatedStatusID = SP.UI.Status.addStatus("Attention: ", message, true);
    SP.UI.Status.setStatusPriColor(strUpdatedStatusID, bgColor);

}

function IsOnPage(pageName) {
    if (window.location.href.toLowerCase().indexOf(pageName.toLowerCase()) > -1) {
        return true;
    } else {
        return false;
    }
}

function getSiteSharingStatus() {   
    if (hostUrl == subSiteUrl) {
        var clientContext = SP.ClientContext.get_current();
        var web = clientContext.get_web();
        var props = web.get_allProperties();
        clientContext.load(web);
        clientContext.load(props);

        clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
            sharingStatus = props.get_item('_site_props_SharingCapability');
            setUI();

        }));

        return sharingStatus;
    }
    else {
        var clientContext = SP.ClientContext.get_current();
        var site = clientContext.get_site();
        var web = site.get_rootWeb();
        var props = web.get_allProperties();
        clientContext.load(web);
        clientContext.load(props);

        clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
            sharingStatus = props.get_item('_site_props_SharingCapability');
            //var sharingStatus1 = props.get_item("_site_props_MainOwner");
            //var sharingStatus2 = props.get_item("_site_props_Type");
            //var sharingStatus3 = props.get_item("_site_props_CreationDate");
            //var sharingStatus4 = props.get_item("_site_props_BlockDate");
            //var sharingStatus5 = props.get_item("_site_props_Compliance");
            //var sharingStatus6 = props.get_item("_site_props_closed");
            //var sharingStatus7 = props.get_item("_site_props_expirydate");
            //var sharingStatus8 = props.get_item("_site_props_policy");	
			
			
            //setUI(sharingStatus1,sharingStatus2,sharingStatus3,sharingStatus4,sharingStatus5);

        }));

        return sharingStatus;
    }
    
}

function setUI(MainOwner, Type, CreationDate, BlockDate, Compliance, closed, expdate, policy, siteid, curuser, isowner) {
    var url = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl;
    if (url.indexOf("https://flextronics365.sharepoint.com/sites") > -1) {
        var policy_icon = "<div class=\"o365cs-nav-topItem o365cs-rsp-tn-hideIfAffordanceOff\"><div>" +
            "<button type=\"button\" class=\"o365cs-nav-item o365cs-nav-button ms-bgc-tdr-h o365button o365cs-topnavText\" role=\"menuitem\" id=opendg_topnav aria-disabled=\"false\"" +
            "aria-selected=\"false\" aria-label=\"Site Policy Type is " + policy + ", Click to change the site policy type\">" +
            "<span class=\"o365cs-topnavText owaimg ms-Icon--policy ms-icon-font-size-20\" aria-hidden=\"true\" id=opendg_topnavspan> </span>&nbsp;&nbsp;" + policy + "&nbsp;&nbsp;" +
            "<div class=\"o365cs-flexPane-unseenitems\"> <span class=\"o365cs-flexPane-unseenCount ms-bgc-tdr ms-fcl-w\" style=\"display: none;\"></span>" +
            "<span class=\"o365cs-flexPane-unseenCount owaimg ms-Icon--starburst ms-icon-font-size-12 ms-bgc-tdr ms-fcl-w\" style=\"display: none;\"> </span> </div></button></div></div>";

        var message = "";
        var curDate = new Date();
        curDate.setHours(0, 0, 0, 0);
        var eDate = new Date(expdate);
        eDate.setHours(0, 0, 0, 0);
        var diffDays = parseInt((eDate - curDate) / (1000 * 60 * 60 * 24));
        setTimeout(function () {
            $(".o365cs-nav-rightMenus").find("div:first").prepend(policy_icon);
            if (policy == "Public") {
                $("#opendg_topnavspan").html("&nbsp;&nbsp;<img src=\"https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/images/navicon/public.png\">");
                //document.styleSheets[0].addRule("span#opendg_topnavspan::before", "content: '\\e155'");
            }
            else {
                $("#opendg_topnavspan").html("&nbsp;&nbsp;<img src=\"https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/images/navicon/private.png\">");
                //document.styleSheets[0].addRule("span#opendg_topnavspan::before", "content: '\\e008'");
            }
            $("#opendg_topnav").on("click", function () {
                var url = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl;
                var username = _spPageContextInfo.userLoginName;
                openDialog("https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + MainOwner + "&type=" + Type + "&cdate=" + CreationDate + "&bdate=" + BlockDate + "&comply=" + Compliance + "&curl=" + currentUrl + "&close=" + closed + "&edate=" + expdate + "&policy=" + policy + "&siteid=" + siteid + "&curuser=" + curuser + "&srclink=nav&isowner=" + isowner);
            });
        }, 1000);
    }
    if (policy == 'High Business Impact') {
        //alert("Sharing: " + sharingStatus);
        if (sharingStatus == "true" || sharingStatus == "True") {
            message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>High Business Impact</b> and <b>Partner Sharing is enabled</b></a></font>";
        }
        else {
            message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>High Business Impact</b></a></font>";
        }
        statusBarBackground = "#f0f0f0";
    }
    else if (policy == 'Medium Business Impact') {
        if (sharingStatus == "true" || sharingStatus == "True") {
            message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>Medium Business Impact</b> and <b>Partner Sharing is enabled</b></a></font>";
        }
        else {
            message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>Medium Business Impact</b></a></font>";
        }
        statusBarBackground = "#f0f0f0";
    }
    else if (policy == 'Low Business Impact') {
        if (sharingStatus == "true" || sharingStatus == "True") {
            message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>Low Business Impact</b> and <b>Partner Sharing is enabled</b></a></font>";
        }
        else {
            message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>Low Business Impact</b></a></font>";
        }
        statusBarBackground = "#f0f0f0";
    }    
    else if (diffDays < 0) {
        if (sharingStatus == "true" || sharingStatus == "True") {
            message = "<font color='#000000'>Information on this site has not yet been classified.  Click <a href='#' id='opendg'>here</a> to set the Policy. <b>Partner sharing is enabled</b></font>";
        }
        else {
            message = "<font color='#000000'>This site looks already expired.  Click <a href='#' id='opendgrtn'>here</a> to extend the site expiry date and to keep this site open.</font>";
        }
        statusBarBackground = "red";
    }
    else if (diffDays >= 0 && diffDays <= 30) {
        if (sharingStatus == "true" || sharingStatus == "True") {
            message = "<font color='#000000'>Information on this site has not yet been classified.  Click <a href='#' id='opendgrtn'>here</a> to set the Policy. <b>Partner sharing is enabled</b></font>";
        }
        else {
            message = "<font color='#000000'>This site will be expired with in " + diffDays + " days .  Click <a href='#' id='opendgrtn'>here</a> to extend the site expiry date.</font>";
        }
        statusBarBackground = "red";
    }   
    else if (Compliance == "false") {
        if (sharingStatus == "true" || sharingStatus == "True") {
            message = "<font color='#000000'>Information on this site has not yet been classified.  Click <a href='#' id='opendg'>here</a> to set the Policy. <b>Partner sharing is enabled</b></font>";
        }
        else {
            message = "<font color='#000000'>Site is not under compliance based on the SharePoint site governance.  Click <a href='#' id='opendg'>here</a> to validate the site compliance.</font>";
        }
        statusBarBackground = "yellow";

    }
   
    if (closed == "true") {
        window.location.href = "https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + MainOwner + "&type=" + Type + "&cdate=" + CreationDate + "&bdate=" + BlockDate + "&comply=" + Compliance + "&curl=" + currentUrl + "&close=" + closed + "&edate=" + expdate + "&policy=" + policy + "&siteid=" + siteid+"&curuser="+curuser;
    }
    // add code to set a policy (reminder red) This sub-site does not have a policy set. Click here to set
    if (message.length > 0) {
        SetStatusBar(message, statusBarBackground);
    }
    
    loadPopup = getParameterByName("fromsource");
    if ((loadPopup != "undefined" && loadPopup != "" && loadPopup != "null" && loadPopup == "compemail"))
    {
        var url = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl;
        var username = _spPageContextInfo.userLoginName;
        openDialog("https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + MainOwner + "&type=" + Type + "&cdate=" + CreationDate + "&bdate=" + BlockDate + "&comply=" + Compliance + "&curl=" + currentUrl + "&close=" + closed + "&edate=" + expdate + "&policy=" + policy + "&siteid=" + siteid + "&curuser=" + curuser + "&srclink=compemail&isowner=" + isowner);
    }

    if ((loadPopup != "undefined" && loadPopup != "" && loadPopup != "null" && loadPopup == "rtnemail"))
    {
        var url = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl;
        var username = _spPageContextInfo.userLoginName;
        openDialog("https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + MainOwner + "&type=" + Type + "&cdate=" + CreationDate + "&bdate=" + BlockDate + "&comply=" + Compliance + "&curl=" + currentUrl + "&close=" + closed + "&edate=" + expdate + "&policy=" + policy + "&siteid=" + siteid + "&curuser=" + curuser + "&srclink=rtnemail&isowner=" + isowner);
    }
   
	$("#opendg").on("click",function() {
	var url = window.location.protocol + "//" + window.location.host +_spPageContextInfo.webServerRelativeUrl;
	var username = _spPageContextInfo.userLoginName;
	openDialog("https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + MainOwner + "&type=" + Type + "&cdate=" + CreationDate + "&bdate=" + BlockDate + "&comply=" + Compliance + "&curl=" + currentUrl + "&close=" + closed + "&edate=" + expdate + "&policy=" + policy + "&siteid=" + siteid + "&curuser=" + curuser + "&srclink=link&isowner=" + isowner);
    });
    
    $("#opendgrtn").on("click",function() {
        var url = window.location.protocol + "//" + window.location.host +_spPageContextInfo.webServerRelativeUrl;
        var username = _spPageContextInfo.userLoginName;
        openDialog("https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + MainOwner + "&type=" + Type + "&cdate=" + CreationDate + "&bdate=" + BlockDate + "&comply=" + Compliance + "&curl=" + currentUrl + "&close=" + closed + "&edate=" + expdate + "&policy=" + policy + "&siteid=" + siteid + "&curuser=" + curuser + "&srclink=rtnemail&isowner=" + isowner);
        });
	function openDialog(pageUrl) { 
	    var w = window.innerWidth;
        var h = window.innerHeight;
        w=w/1.6;
        h=h/1.2;
        
        var options = {
            url: pageUrl,
            title: 'Site Governance',
            allowMaximize: true,
            showClose: true,
            width: w,
            height: h
        };
	    SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', options);       
	}
}

function getParameterByName(name) {
    try
    {
        var params = document.URL.split("?")[1].split("&");
        var strParams = "";
        for (var i = 0; i < params.length; i = i + 1) {
            var singleParam = params[i].split("=");
            if (singleParam[0] == name)
                return singleParam[1];
        }
    }
    catch(err)
    {
        return "null";
    }
}

function loadScript(url, callback) {
    var head = document.getElementsByTagName("head")[0];
    var script = document.createElement("script");
    script.src = url;

    // Attach handlers for all browsers
    var done = false;
    script.onload = script.onreadystatechange = function () {
        if (!done && (!this.readyState
					|| this.readyState == "loaded"
					|| this.readyState == "complete")) {
            done = true;

            // Continue your code
            callback();

            // Handle memory leak in IE
            script.onload = script.onreadystatechange = null;
            head.removeChild(script);
        }
    };

    head.appendChild(script);
}