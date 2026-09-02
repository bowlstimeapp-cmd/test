(function () {

    "use strict";


    /*
    ============================================================
    CONFIGURATION
    ============================================================
    */


    /*
    ============================================================
    SELECTED PASS HOLDER
    ============================================================
    */

    var selectedPassHolder = null;

    var passHolderSearchTimer = null;


    /*
    ============================================================
    ATTACHMENT CONFIGURATION
    ============================================================
    */

    var MAX_ATTACHMENT_SIZE =
        10 * 1024 * 1024;

    var ALLOWED_ATTACHMENT_TYPES = [
        "image/jpeg",
        "image/png",
         "application/pdf"
    ];

    var ALLOWED_ATTACHMENT_EXTENSIONS = [
        ".jpg",
        ".jpeg",
        ".png",
        ".pdf"
    ];


    /*
    ============================================================
    ATTACHMENT HELPER FUNCTIONS

    These are used throughout the attachment upload code
    (file input validation, size display, SharePoint filename
    handling, and error messages).
    ============================================================
    */

    function formatFileSize(bytes) {

        if (
            bytes === null ||
            bytes === undefined ||
            isNaN(bytes)
        ) {

            return "";

        }


        if (bytes < 1024) {

            return bytes + " B";

        }


        if (bytes < 1024 * 1024) {

            return (
                bytes / 1024
            ).toFixed(1) + " KB";

        }


        return (
            bytes / (1024 * 1024)
        ).toFixed(1) + " MB";

    }


    function getAttachmentExtension(fileName) {

        var safeName =
            String(
                fileName || ""
            );


        var lastDot =
            safeName.lastIndexOf(".");


        if (lastDot === -1) {

            return "";

        }


        return safeName.substring(
            lastDot
        ).toLowerCase();

    }


    function validateAttachmentFile(file) {

        if (!file) {

            return {

                valid: false,

                message:
                    "No file was selected."

            };

        }


        if (file.size > MAX_ATTACHMENT_SIZE) {

            return {

                valid: false,

                message:
                    "File is too large. Maximum file size is 10 MB."

            };

        }


        var extension =
            getAttachmentExtension(
                file.name
            );


        var typeAllowed =
            ALLOWED_ATTACHMENT_TYPES.indexOf(
                file.type
            ) !== -1;


        var extensionAllowed =
            ALLOWED_ATTACHMENT_EXTENSIONS.indexOf(
                extension
            ) !== -1;


        /*
        ------------------------------------------------------------
        Some browsers/OS combinations report an empty or generic
        MIME type for certain files, so we accept the file if
        EITHER the MIME type OR the file extension is allowed.
        ------------------------------------------------------------
        */

        if (
            !typeAllowed &&
            !extensionAllowed
        ) {

            return {

                valid: false,

                message:
                    "Unsupported file type. Allowed types: JPG, PNG, PDF."

            };

        }


        return {

            valid: true,

            message: ""

        };

    }


    function showAttachmentError(message) {

        var errorContainer =
            document.getElementById(
                "passAttachmentError"
            );


        if (!errorContainer) {

            return;

        }


        errorContainer.textContent =
            message ||
            "";


        errorContainer.style.display =
            message
                ? "block"
                : "none";

    }


    /*
    ============================================================
    SHAREPOINT LIST CONFIGURATION
    ============================================================
    */


    var SITE_URL =
        "https://mcga.sharepoint.com/sites/InformationAssurance";


    var LIST_NAME =
        "Building_Pass_Request_Public";


    var LIST_API_URL =
        SITE_URL +
        "/_api/web/lists/getbytitle('" +
        LIST_NAME +
        "')/items";


    var LIST_ITEM_TYPE =
        "SP.Data.Building_x005f_Pass_x005f_RequestListItem";


    /*
    ============================================================
    ADMIN USERS
    ============================================================

    Only these SharePoint accounts will see/access the
    Admin Dashboard.

    ============================================================
    */

    var ADMIN_USERS = [

        "joshua.davis@mcga.gov.uk",

        "chris.townsend@mcga.gov.uk",

        "hqsecuritymanager@mcga.gov.uk", 

        "frontdesksecurity@mcga.gov.uk"

    ];


    /*
    ============================================================
    SITE ACCESS PERMISSIONS
    ============================================================

    EDIT THIS SECTION TO DEFINE THE ACCESS PERMISSIONS
    AVAILABLE FOR EACH SITE.

    The user can select MULTIPLE permissions.

    The selected permissions will be saved to SharePoint
    in the existing comma-separated Access field.

    ============================================================
    */

    var SITE_ACCESS_OPTIONS = {

        "Aberdeen": [
            "Main Office",
            "Data Centre",
        ],

        "Belfast": [
            "Main Office",
            "IT Secure Room",
 
        ],

        "Beverley": [
            "N/A - No Electronic Pass"
        ],

        "Cardiff": [
            "N/A - No Electronic Pass"
        ],

        "Colchester": [
            "N/A - No Electronic Pass"
        ],

        "Daedalus": [
            "Electric Pass"
        ],

        "Dover": [
            "Main Entrance",
            "Office Area",
            "Restricted Area"
        ],

        "Falmouth": [
            "N/A - No Electronic Pass"
        ],

        "Glasgow": [
            "N/A - No Electronic Pass"
        ],

        "Holyhead": [
            "N/A - No Electronic Pass"
        ],

        "Humber": [
            "N/A - No Electronic Pass"
        ],

        "JRCC": [
            "Group 4 (Visitor/Contractor) - General Access",
            "Group 5 (ICT) - General Access, IT Area, Training Area, Operations Room, Planning Room",
            "Group 6 (ICT+) - General Access, IT Area, Training Area, Operations Room, Planning Room & Data Hall",
            "Group 7 (Operations) - General Access, Training Area & Operations Room",
            "Group 8 (Operations+) - General Access, Training Area, Operations Room & Planning Room",
            "Group 9 (Trainee) - General Access & Training Area",
            "Group 10 (FM Contractor) - General, IT Area, Training Area, Operations Room, FM Areas & Planning Room",
            "Group 11 (FM Contractor +) - General, IT Area, Training Area, Operations Room, FM Area, Planning Room & Data Hall"
        ],

        "Liverpool": [
            "N/A = No Electronic Pass",
        ],

        "London": [
            "N/A = No Electronic Pass",
        ],

        "Milford Haven": [
            "N/A = No Electronic Pass",
        ],

        "Plymouth": [
            "N/A = No Electronic Pass",
        ],

        "Shetland": [
            "N/A = No Electronic Pass",
        ],

        "Spring Place": [
            "MCA General Access",
            "MCA Marine Office",
            "MCA IT Build Room",
            "MCA Disaster Recovery Suite",
            "MCA Data Centre",
            "MCA RCIT Interview Room"
        ],

        "Stornoway": [
            "N/A = No Electronic Pass",
        ],

        "Swansea": [
            "N/A = No Electronic Pass",
        ],

        "Torbay": [
            "N/A = No Electronic Pass",
        ],

        "Tyneside": [
            "N/A = No Electronic Pass",
        ]

    };


    /*
    ============================================================
    APPLICATION VIEW
    ============================================================
    */

    var currentView = "home";


    /*
    ============================================================
    DASHBOARD CONTAINER
    ============================================================
    */

    var dashboard =
        document.getElementById("passDashboard");


    if (!dashboard) {

        console.error(
            "Pass Dashboard: #passDashboard not found."
        );

        return;

    }


    /*
    ============================================================
    DATA
    ============================================================
    */

    var allData = [];

    var data = [];

    var currentSearch = "";

    var currentSite = "ALL";

    var currentStatus = "ALL";


    /*
    ============================================================
    CURRENT LOGGED-IN USER
    ============================================================
    */

    var currentUserEmail = "";

    var currentUserDisplayName = "";

    var isAdmin = false;


    /*
    ============================================================
    AVAILABLE SITES
    ============================================================
    */

    var availableSites = [];


    /*
    ============================================================
    INITIAL LOADING MESSAGE
    ============================================================
    */

    dashboard.innerHTML = `

        <div style="
            padding:30px;
            background:#f3f2f1;
            border-radius:10px;
            font-family:'Segoe UI',Arial,sans-serif;
            color:#323130;
            text-align:center;
        ">

            <div style="
                font-size:20px;
                font-weight:600;
                margin-bottom:8px;
            ">
                Loading Building Pass Request Dashboard...
            </div>

            <div style="
                font-size:13px;
                color:#605e5c;
            ">
                Retrieving the latest pass data.
            </div>

        </div>

    `;


    /*
    ============================================================
    HELPERS
    ============================================================
    */

    function escapeHtml(value) {

        if (
            value === null ||
            value === undefined
        ) {

            return "";

        }


        return String(value)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");

    }


    function normalise(value) {

        return String(
            value || ""
        )
        .trim()
        .toLowerCase();

    }


    function isAdminUser() {

        var email =
            normalise(
                currentUserEmail
            );


        return ADMIN_USERS.some(
            function (adminEmail) {

                return normalise(
                    adminEmail
                ) === email;

            }
        );

    }


    function formatDate(value) {

        if (!value) {

            return "—";

        }


        var date =
            new Date(value);


        if (
            isNaN(
                date.getTime()
            )
        ) {

            return escapeHtml(value);

        }


        return date.toLocaleDateString(
            "en-GB",
            {
                day: "2-digit",
                month: "short",
                year: "numeric"
            }
        );

    }


    function formatDateTime(value) {

        if (!value) {

            return "—";

        }


        var date =
            new Date(value);


        if (
            isNaN(
                date.getTime()
            )
        ) {

            return escapeHtml(value);

        }


        return (
            date.toLocaleDateString(
                "en-GB",
                {
                    day: "2-digit",
                    month: "short",
                    year: "numeric"
                }
            )
            +
            " "
            +
            date.toLocaleTimeString(
                "en-GB",
                {
                    hour: "2-digit",
                    minute: "2-digit"
                }
            )
        );

    }


    /*
    ============================================================
    GET ACCESS OPTIONS FOR SITE
    ============================================================
    */

    function getAccessOptionsForSite(site) {

        if (
            !site ||
            !SITE_ACCESS_OPTIONS[site]
        ) {

            return [];

        }


        return SITE_ACCESS_OPTIONS[site];

    }


    /*
    ============================================================
    GET SELECTED ACCESS PERMISSIONS
    ============================================================
    */

    function getSelectedAccessPermissions() {

        var accessSelect =
            document.getElementById(
                "formAccess"
            );


        if (!accessSelect) {

            return [];

        }


        return Array.from(
            accessSelect.selectedOptions
        )
        .map(
            function (option) {

                return option.value;

            }
        )
        .filter(
            function (value) {

                return value !== "";

            }
        );

    }


    /*
    ============================================================
    UPDATE ACCESS PERMISSIONS DROPDOWN
    ============================================================
    */

    function updateAccessPermissionsDropdown(
        clearSelection
    ) {

        var siteSelect =
            document.getElementById(
                "formSite"
            );


        var accessSelect =
            document.getElementById(
                "formAccess"
            );


        var accessHelp =
            document.getElementById(
                "passAccessStatus"
            );


        if (
            !siteSelect ||
            !accessSelect
        ) {

            return;

        }


        var selectedSite =
            siteSelect.value;


        var options =
            getAccessOptionsForSite(
                selectedSite
            );


        if (clearSelection !== false) {

            accessSelect.innerHTML = "";

        }


        if (!selectedSite) {

            accessSelect.disabled =
                true;


            accessSelect.innerHTML = `

                <option value="">
                    Select a site first
                </option>

            `;


            if (accessHelp) {

                accessHelp.className =
                    "pass-site-loading";

                accessHelp.textContent =
                    "Select a site to see available access areas.";

            }


            return;

        }


        accessSelect.disabled =
            false;


        accessSelect.innerHTML = `

            <option value="">
                Select access areas...
            </option>

        `;


        options.forEach(
            function (optionValue) {

                var option =
                    document.createElement(
                        "option"
                    );


                option.value =
                    optionValue;


                option.textContent =
                    optionValue;


                accessSelect.appendChild(
                    option
                );

            }
        );


        if (options.length === 0) {

            accessSelect.disabled =
                true;


            accessSelect.innerHTML = `

                <option value="">
                    No access areas configured
                </option>

            `;


            if (accessHelp) {

                accessHelp.className =
                    "pass-site-error";

                accessHelp.textContent =
                    "No access permissions have been configured for this site.";

            }

            return;

        }


        if (accessHelp) {

            accessHelp.className =
                "pass-site-loading";

            accessHelp.textContent =
                "Hold Ctrl (Windows) or Command (Mac) to select multiple areas.";

        }

    }


    /*
    ============================================================
    GET CURRENT SHAREPOINT USER
    ============================================================
    */

    function getCurrentUser() {

        return fetch(
            SITE_URL +
            "/_api/web/currentuser",
            {
                method: "GET",

                credentials: "same-origin",

                headers: {
                    "Accept":
                        "application/json;odata=verbose"
                }
            }
        )
        .then(
            function (response) {

                if (!response.ok) {

                    throw new Error(
                        "Unable to determine the current SharePoint user. HTTP " +
                        response.status
                    );

                }


                return response.json();

            }
        )
        .then(
            function (json) {

                var user =
                    json.d;


                currentUserEmail =
                    String(
                        user.Email || ""
                    )
                    .trim()
                    .toLowerCase();


                currentUserDisplayName =
                    String(
                        user.Title || ""
                    )
                    .trim();


                isAdmin =
                    isAdminUser();


                console.log(
                    "Current SharePoint user:",
                    currentUserDisplayName,
                    currentUserEmail
                );


                console.log(
                    "Admin access:",
                    isAdmin
                );


                return user;

            }
        );

    }


    /*
    ============================================================
    GET AVAILABLE SITES
    ============================================================
    */

    function getAvailableSites() {

        availableSites =
            Object.keys(
                SITE_ACCESS_OPTIONS
            )
            .sort(
                function (a, b) {

                    return a.localeCompare(
                        b,
                        "en-GB",
                        {
                            sensitivity:
                                "base"
                        }
                    );

                }
            );


        console.log(
            "Available sites from SITE_ACCESS_OPTIONS:",
            availableSites
        );


        return Promise.resolve(
            availableSites
        );

    }


   /* ============================================================
   STATUS
   ============================================================ */

function getStatus(record) {

    var value = String(record.Status || "").trim();
    var lowerValue = value.toLowerCase();

    // ------------------------------------------------------------
    // High-level dashboard status
    // ------------------------------------------------------------

    if (lowerValue.indexOf("pass issued") !== -1) {

        return {
            label: "Issued",
            description: value || "Pass issued",
            className: "status-green"
        };

    }

    if (lowerValue.indexOf("rejected") !== -1) {

        return {
            label: "Rejected",
            description: value || "Application rejected",
            className: "status-red"
        };

    }

    if (lowerValue.indexOf("pass with") !== -1) {

        return {
            label: "Pending",
            description: value || "Pass in progress",
            className: "status-amber"
        };

    }

    // ------------------------------------------------------------
    // Other
    // ------------------------------------------------------------

    return {
        label: "Other",
        description: value || "Status unavailable",
        className: "status-grey"
    };
}


    /*
    ============================================================
    EXPIRY
    ============================================================
    */

    function getExpiry(record) {

        if (!record.Expiry_Date) {

            return {

                label: "No expiry",

                className: "expiry-grey"

            };

        }


        var expiry =
            new Date(
                record.Expiry_Date
            );


        var today =
            new Date();


        expiry.setHours(
            23,
            59,
            59,
            999
        );


        today.setHours(
            0,
            0,
            0,
            0
        );


        var days =
            Math.ceil(
                (
                    expiry.getTime()
                    -
                    today.getTime()
                )
                /
                (
                    1000 *
                    60 *
                    60 *
                    24
                )
            );


        if (days < 0) {

            return {

                label: "Expired",

                className: "expiry-red"

            };

        }


        if (days <= 90) {

            return {

                label:
                    days +
                    " days remaining",

                className:
                    "expiry-amber"

            };

        }


        return {

            label: "Valid",

            className:
                "expiry-green"

        };

    }


    /*
    ============================================================
    ACCESS TAGS
    ============================================================
    */

    function accessTags(value) {

        if (!value) {

            return "—";

        }


        return String(value)
            .split(",")
            .map(
                function (item) {

                    return (
                        '<span class="access-tag">' +
                        escapeHtml(
                            item.trim()
                        ) +
                        '</span>'
                    );

                }
            )
            .join("");

    }


    /*
    ============================================================
    GET SITES
    ============================================================
    */

    function getSites() {

        var sites = {};


        data.forEach(
            function (record) {

                if (record.Site) {

                    sites[
                        record.Site
                    ] = true;

                }

            }
        );


        return Object.keys(sites)
            .sort();

    }


    /*
    ============================================================
    FILTER DATA
    ============================================================
    */

    function getFilteredData() {

        return data.filter(
            function (record) {

                var searchText =
                    normalise(
                        currentSearch
                    );


                var matchesSearch =
                    !searchText

                    ||

                    String(
                        record.ID || ""
                    )
                    .toLowerCase()
                    .indexOf(
                        searchText
                    ) !== -1

                    ||

                    String(
                        record.Pass_Holder || ""
                    )
                    .toLowerCase()
                    .indexOf(
                        searchText
                    ) !== -1

                    ||

                    String(
                        record.Site || ""
                    )
                    .toLowerCase()
                    .indexOf(
                        searchText
                    ) !== -1

                    ||

                    String(
                        record.Clearance_Level || ""
                    )
                    .toLowerCase()
                    .indexOf(
                        searchText
                    ) !== -1

                    ||

                    String(
                        record.Pass_Requester || ""
                    )
                    .toLowerCase()
                    .indexOf(
                        searchText
                    ) !== -1;


                var status =
                    getStatus(record);


                var matchesStatus =
                    currentStatus === "ALL"
                    ||
                    status.label ===
                    currentStatus;


                var matchesSite =
                    currentSite === "ALL"
                    ||
                    record.Site ===
                    currentSite;


                return (
                    matchesSearch
                    &&
                    matchesStatus
                    &&
                    matchesSite
                );

            }
        );

    }


    /*
    ============================================================
    COUNTS
    ============================================================
    */

    function getCounts() {

        var issued = 0;

        var pending = 0;

        var rejected = 0;

        var expired = 0;


        data.forEach(
            function (record) {

                var status =
                    getStatus(record);


                if (
                    status.label ===
                    "Issued"
                ) {

                    issued++;

                }


                if (
                    status.label ===
                    "Pending"
                ) {

                    pending++;

                }


                if (
                    status.label ===
                    "Rejected"
                ) {

                    rejected++;

                }


                var expiry =
                    getExpiry(record);


                if (
                    expiry.label ===
                    "Expired"
                ) {

                    expired++;

                }

            }
        );


        return {

            issued: issued,

            pending: pending,

            rejected: rejected,

            expired: expired

        };

    }


    /*
    ============================================================
    CSS
    ============================================================
    */

    var style =
        document.createElement(
            "style"
        );


    style.textContent = `

        #passDashboard {
            font-family:
                "Segoe UI",
                Arial,
                sans-serif;

            color:#323130;

            width:100%;

            max-width:1250px;

            margin:0 auto;
        }


        #passDashboard * {
            box-sizing:border-box;
        }


        .dashboard-header {

            background:
                linear-gradient(
                    135deg,
                    #003b5c,
                    #0078d4
                );

            color:white;

            padding:28px 30px;

            border-radius:10px;

            margin-bottom:18px;
        }


        .dashboard-header h1 {

            margin:0;

            font-size:28px;

            font-weight:600;
        }

/* ============================================================
   ISSUE PASS BUTTON
   ============================================================ */

.action-cell {
    white-space: nowrap;
}


.issue-pass-button {
    margin-top:5%;
    appearance: none;
    border: 1px solid #198754;
    background: #198754;
    color: #ffffff;
    white-space:nowrap;
    margin-left: auto;
    display: block;

    padding: 6px 12px;

    border-radius: 5px;

    font-family: inherit;
    font-size: 13px;
    font-weight: 600;

    line-height: 1.4;

    cursor: pointer;

    transition:
        background-color 0.15s ease,
        border-color 0.15s ease,
        box-shadow 0.15s ease,
        transform 0.05s ease;
}


.issue-pass-button:hover {
    background: #157347;
    border-color: #146c43;
}


.issue-pass-button:focus {
    outline: none;

    box-shadow:
        0 0 0 3px rgba(25, 135, 84, 0.20);
}


.issue-pass-button:active {
    background: #146c43;
    border-color: #13653f;

    transform: translateY(1px);
}


.issue-pass-button:disabled {
    opacity: 0.6;
    cursor: not-allowed;
}

        .dashboard-header p {

            margin:
                6px 0 0 0;

            opacity:.85;

            font-size:14px;
        }


        .admin-banner {

            background:#fff4ce;

            color:#8a6d00;

            border-left:
                4px solid #ffb900;

            padding:10px 14px;

            border-radius:6px;

            margin-bottom:15px;

            font-size:13px;

            font-weight:600;
        }


        .summary-grid {

            display:grid;

            grid-template-columns:
                repeat(5,1fr);

            gap:14px;

            margin-bottom:18px;
        }


        .summary-card {

            background:white;

            border:
                1px solid #edebe9;

            border-radius:9px;

            padding:18px;

            box-shadow:
                0 2px 6px
                rgba(0,0,0,.05);
        }


        .summary-label {

            font-size:11px;

            font-weight:600;

            text-transform:uppercase;

            color:#605e5c;

            letter-spacing:.5px;
        }


        .summary-number {

            font-size:30px;

            font-weight:600;

            margin-top:5px;
        }


        .number-blue {
            color:#0078d4;
        }


        .number-green {
            color:#107c10;
        }


        .number-amber {
            color:#ca5010;
        }


        .number-red {
            color:#a4262c;
        }


        .controls {

            display:flex;

            gap:12px;

            align-items:center;

            background:white;

            border:
                1px solid #edebe9;

            padding:15px;

            border-radius:9px;

            margin-bottom:15px;

            box-shadow:
                0 2px 6px
                rgba(0,0,0,.05);
        }


        .search-box {

            flex:1;

            min-width:200px;

            padding:
                10px 13px;

            border:
                1px solid #8a8886;

            border-radius:5px;

            font-family:
                "Segoe UI",
                Arial,
                sans-serif;

            font-size:14px;
        }


        .search-box:focus {

            outline:none;

            border-color:#0078d4;

            box-shadow:
                0 0 0 1px #0078d4;
        }


        .filter-select {

            padding:
                10px 13px;

            border:
                1px solid #8a8886;

            border-radius:5px;

            background:white;

            font-family:
                "Segoe UI",
                Arial,
                sans-serif;

            font-size:14px;

            min-width:150px;
        }


        .record-count {

            color:#605e5c;

            font-size:13px;

            white-space:nowrap;
        }


        .table-container {

            background:white;

            border:
                1px solid #edebe9;

            border-radius:9px;

            overflow:hidden;

            box-shadow:
                0 2px 6px
                rgba(0,0,0,.05);
        }


        .pass-table {

            width:100%;

            border-collapse:collapse;
        }


        .pass-table th {

            background:#f3f2f1;

            color:#323130;

            font-size:12px;

            font-weight:600;

            text-align:left;

            padding:13px 14px;

            border-bottom:
                1px solid #edebe9;

            white-space:nowrap;
        }


        .pass-table td {

            padding:14px;

            border-bottom:
                1px solid #f3f2f1;

            font-size:13px;

            vertical-align:middle;
        }


        .pass-table tbody tr {

            cursor:pointer;

            transition:
                background .15s;
        }


        .pass-table tbody tr:hover {

            background:#f5f9fc;
        }


        .pass-table tbody tr:last-child td {

            border-bottom:none;
        }


        .id-cell {

            font-weight:600;

            color:#0078d4;
        }


        .holder-cell {

            font-weight:600;
        }


        .status-pill {

            display:inline-block;

            padding:5px 10px;

            border-radius:20px;

            font-size:11px;

            font-weight:600;

            white-space:normal;
            overflow-wrap: break-word;
            word-wrap: break-word;
            line-height: 1.4;
        }


        .status-green {

            background:#dff6dd;

            color:#107c10;
        }


        .status-amber {

            background:#fff4ce;

            color:#8a6d00;
        }


        .status-red {

            background:#fde7e9;

            color:#a4262c;
        }


        .status-grey {

            background:#edebe9;

            color:#605e5c;
        }


        .expiry-green {

            color:#107c10;

            font-weight:600;

            font-size:11px;
        }


        .expiry-amber {

            color:#ca5010;

            font-weight:600;

            font-size:11px;
        }


        .expiry-red {

            color:#a4262c;

            font-weight:600;

            font-size:11px;
        }


        .expiry-grey {

            color:#605e5c;

            font-size:11px;
        }


        .empty-state {

            padding:50px 20px;

            text-align:center;

            color:#605e5c;
        }


        .error-state {

            background:#fde7e9;

            border-left:
                5px solid #a4262c;

            padding:25px;

            border-radius:8px;

            font-family:
                "Segoe UI",
                Arial,
                sans-serif;
        }


        .error-state h2 {

            margin:
                0 0 8px 0;

            color:#a4262c;

            font-size:20px;
        }


        .error-state p {

            margin:5px 0;

            color:#323130;

            font-size:14px;
        }


        .details-overlay {

            position:fixed;

            top:0;

            left:0;

            right:0;

            bottom:0;

            background:
                rgba(0,0,0,.45);

            z-index:99999;

            display:flex;

            justify-content:center;

            align-items:center;

            padding:20px;
        }


        .details-modal {

            background:white;

            border-radius:10px;

            width:100%;

            max-width:900px;

            max-height:90vh;

            overflow:auto;

            box-shadow:
                0 15px 50px
                rgba(0,0,0,.3);
        }


        .modal-header {

            display:flex;

            justify-content:space-between;

            align-items:center;

            background:
                linear-gradient(
                    135deg,
                    #003b5c,
                    #0078d4
                );

            color:white;

            padding:20px 24px;
        }


        .modal-header h2 {

            margin:0;

            font-size:21px;
        }


        .close-button {

            border:none;

            background:
                rgba(255,255,255,.15);

            color:white;

            width:34px;

            height:34px;

            border-radius:50%;

            font-size:20px;

            cursor:pointer;
        }


        .close-button:hover {

            background:
                rgba(255,255,255,.3);
        }


        .modal-body {

            padding:24px;
        }


        .detail-grid {

            display:grid;

            grid-template-columns:
                1fr 1fr;

            gap:20px;
        }


        .detail-section {

            border:
                1px solid #edebe9;

            border-radius:8px;

            overflow:hidden;
        }


        .detail-section-full {

            grid-column:
                1 / -1;
        }


        .detail-section-title {

            background:#f3f2f1;

            padding:12px 15px;

            font-weight:600;

            font-size:14px;
        }


        .detail-section-body {

            padding:15px;
        }


        .detail-row {

            display:grid;

            grid-template-columns:
                145px 1fr;

            gap:12px;

            padding:8px 0;

            border-bottom:
                1px solid #f3f2f1;
        }


        .detail-row:last-child {

            border-bottom:none;
        }


        .detail-label {

            color:#605e5c;

            font-size:12px;

            font-weight:600;
        }


        .detail-value {

            font-size:13px;

            word-break:break-word;
        }


        .access-tags {

            line-height:1.8;
        }


        .access-tag {

            display:inline-block;

            background:#e8f1f8;

            color:#005a9e;

            padding:5px 9px;

            margin:3px;

            border-radius:4px;

            font-size:11px;
        }


        .justification {

            background:#f8f9fa;

            border-left:
                4px solid #0078d4;

            padding:13px;

            line-height:1.5;

            font-size:13px;
        }


        .person {

            display:flex;

            align-items:center;

            gap:12px;

            margin-bottom:10px;
        }


        .person img {

            width:50px;

            height:50px;

            border-radius:50%;

            object-fit:cover;

            background:#edebe9;
        }


        .person-name {

            font-weight:600;

            font-size:14px;
        }


        .person-email {

            color:#605e5c;

            font-size:12px;
        }


        /*
        ========================================================
        APPLICATION NAVIGATION
        ========================================================
        */

        .pass-app-nav {

            display:flex;

            justify-content:space-between;

            align-items:center;

            gap:15px;

            margin-bottom:18px;

            background:white;

            border:
                1px solid #edebe9;

            border-radius:9px;

            padding:12px 15px;

            box-shadow:
                0 2px 6px
                rgba(0,0,0,.05);
        }


        .pass-app-nav-title {

            font-size:14px;

            font-weight:600;

            color:#323130;
        }


        .pass-app-nav-buttons {

            display:flex;

            gap:8px;

            flex-wrap:wrap;
        }


        .pass-nav-button {

            border:none;

            border-radius:5px;

            padding:9px 14px;

            font-family:
                "Segoe UI",
                Arial,
                sans-serif;

            font-size:13px;

            font-weight:600;

            cursor:pointer;

            background:#f3f2f1;

            color:#323130;
        }


        .pass-nav-button:hover {

            background:#e1dfdd;
        }


        .pass-nav-button-primary {

            background:#0078d4;

            color:white;
        }


        .pass-nav-button-primary:hover {

            background:#106ebe;
        }


        .pass-nav-button-admin {

            background:#003b5c;

            color:white;
        }


        .pass-nav-button-admin:hover {

            background:#002b44;
        }


        /*
        ========================================================
        HOME
        ========================================================
        */

        .pass-home {

            font-family:
                "Segoe UI",
                Arial,
                sans-serif;
        }


        .pass-home-header {

            background:
                linear-gradient(
                    135deg,
                    #003b5c,
                    #0078d4
                );

            color:white;

            padding:40px;

            border-radius:10px;

            margin-bottom:20px;
        }


        .pass-home-header h1 {

            margin:0;

            font-size:30px;

            font-weight:600;
        }


        .pass-home-header p {

            margin:8px 0 0 0;

            font-size:15px;

            opacity:.9;
        }


        .pass-choice-grid {

            display:grid;

            grid-template-columns:
                repeat(2, 1fr);

            gap:20px;
        }


        .pass-choice-card {

            background:white;

            border:
                1px solid #edebe9;

            border-radius:10px;

            padding:30px;

            box-shadow:
                0 2px 8px
                rgba(0,0,0,.06);

            cursor:pointer;

            transition:
                transform .15s,
                box-shadow .15s,
                border-color .15s;
        }


        .pass-choice-card:hover {

            transform:
                translateY(-2px);

            border-color:#0078d4;

            box-shadow:
                0 5px 15px
                rgba(0,0,0,.10);
        }


        .pass-choice-icon {

            width:48px;

            height:48px;

            border-radius:50%;

            display:flex;

            align-items:center;

            justify-content:center;

            background:#e8f1f8;

            color:#0078d4;

            font-size:23px;

            margin-bottom:18px;
        }


        .pass-choice-card h2 {

            margin:
                0 0 8px 0;

            font-size:20px;

            color:#323130;
        }


        .pass-choice-card p {

            margin:0;

            color:#605e5c;

            font-size:14px;

            line-height:1.5;
        }


        /*
        ========================================================
        PASS HOLDER USER PICKER
        ========================================================
        */

        .pass-user-picker {
            position: relative;
        }


        .pass-user-results {
            display: none;

            position: absolute;

            top: 100%;
            left: 0;
            right: 0;

            z-index: 1000;

            background: #ffffff;

            border: 1px solid #d0d7de;

            border-radius: 0 0 6px 6px;

            box-shadow:
                0 4px 12px rgba(0, 0, 0, 0.12);

            max-height: 300px;

            overflow-y: auto;
        }


        .pass-user-result {
            display: flex;

            align-items: center;

            gap: 12px;

            padding: 12px 14px;

            cursor: pointer;

            border-bottom: 1px solid #eeeeee;

            transition:
                background-color 0.15s ease;
        }


        .pass-user-result:last-child {
            border-bottom: none;
        }


        .pass-user-result:hover {
            background-color: #f3f6f9;
        }


        .pass-user-result-icon {
            width: 36px;
            height: 36px;

            display: flex;

            align-items: center;
            justify-content: center;

            background: #e8f1fb;

            border-radius: 50%;

            flex-shrink: 0;
        }


        .pass-user-result-details {
            min-width: 0;
        }


        .pass-user-result-name {
            font-weight: 600;

            color: #1f2937;

            font-size: 14px;
        }


        .pass-user-result-email {
            margin-top: 3px;

            color: #6b7280;

            font-size: 13px;

            overflow: hidden;

            text-overflow: ellipsis;

            white-space: nowrap;
        }


        .pass-user-no-results {
            padding: 14px;

            color: #6b7280;

            font-size: 14px;
        }


        .pass-user-search-error {
            padding: 14px;

            color: #b91c1c;

            font-size: 14px;
        }


        /*
        ========================================================
        SELECTED USER
        ========================================================
        */

        .pass-user-selected {
            display: none;

            margin-top: 8px;
        }


        .pass-selected-user {
            display: flex;

            align-items: center;

            gap: 12px;

            padding: 10px 12px;

            background: #f3f7fb;

            border: 1px solid #b8cde3;

            border-radius: 6px;
        }


        .pass-selected-user-icon {
            width: 38px;
            height: 38px;

            display: flex;

            align-items: center;
            justify-content: center;

            background: #dceaf7;

            border-radius: 50%;

            flex-shrink: 0;
        }


        .pass-selected-user-details {
            flex: 1;

            min-width: 0;
        }


        .pass-selected-user-name {
            font-weight: 600;

            color: #1f2937;

            font-size: 14px;
        }


        .pass-selected-user-email {
            margin-top: 3px;

            color: #6b7280;

            font-size: 13px;
        }


        .pass-selected-user-remove {
            width: 30px;
            height: 30px;

            border: none;

            background: transparent;

            color: #6b7280;

            font-size: 22px;

            line-height: 1;

            cursor: pointer;

            border-radius: 4px;
        }


        .pass-selected-user-remove:hover {
            background: #e5e7eb;

            color: #111827;
        }


        /*
        ========================================================
        FORM
        ========================================================
        */

        .pass-form-container {

            background:white;

            border:
                1px solid #edebe9;

            border-radius:10px;

            padding:28px;

            box-shadow:
                0 2px 8px
                rgba(0,0,0,.06);
        }


        .pass-form-header {

            margin-bottom:25px;
        }


        .pass-form-header h1 {

            margin:0;

            font-size:26px;

            color:#003b5c;
        }


        .pass-form-header p {

            margin:
                6px 0 0 0;

            color:#605e5c;

            font-size:14px;
        }


        .pass-form-grid {

            display:grid;

            grid-template-columns:
                1fr 1fr;

            gap:18px;
        }


        .pass-form-field {

            display:flex;

            flex-direction:column;

            gap:6px;
        }


        .pass-form-field-full {

            grid-column:
                1 / -1;
        }


        .pass-form-label {

            font-size:13px;

            font-weight:600;

            color:#323130;
        }


        .pass-form-required {

            color:#a4262c;
        }


        .pass-form-input,
        .pass-form-select,
        .pass-form-textarea {

            width:100%;

            padding:
                10px 12px;

            border:
                1px solid #8a8886;

            border-radius:5px;

            font-family:
                "Segoe UI",
                Arial,
                sans-serif;

            font-size:14px;

            color:#323130;

            background:white;
        }


        /*
        --------------------------------------------------------
        MULTI-SELECT ACCESS DROPDOWN
        --------------------------------------------------------
        */

        .pass-access-select {

            min-height:130px;

            padding:6px 8px;

            cursor:pointer;
        }


        .pass-access-select:disabled {

            background:#f3f2f1;

            color:#a19f9d;

            cursor:not-allowed;
        }


        .pass-access-select option {

            padding:7px 8px;
        }


        .pass-access-select option:checked {

            background:#0078d4;

            color:white;
        }


        .pass-form-input:focus,
        .pass-form-select:focus,
        .pass-form-textarea:focus {

            outline:none;

            border-color:#0078d4;

            box-shadow:
                0 0 0 1px #0078d4;
        }


        .pass-form-textarea {

            min-height:110px;

            resize:vertical;
        }


        .pass-form-help {

            font-size:11px;

            color:#605e5c;
        }


        /*
        ========================================================
        ATTACHMENT UPLOAD
        ========================================================
        */

        .pass-attachment-wrapper {

            width:100%;
        }


        .pass-attachment-input {

            width:100%;

            padding:9px 10px;

            border:
                1px dashed #8a8886;

            border-radius:6px;

            background:#faf9f8;

            font-family:
                "Segoe UI",
                Arial,
                sans-serif;

            font-size:13px;

            color:#323130;

            cursor:pointer;
        }


        .pass-attachment-input:hover {

            border-color:#0078d4;

            background:#f5f9fc;
        }


        .pass-attachment-input:focus {

            outline:none;

            border-color:#0078d4;

            box-shadow:
                0 0 0 1px #0078d4;
        }


        .pass-attachment-selected {

            display:none;

            margin-top:8px;

            padding:9px 11px;

            background:#f3f7fb;

            border:
                1px solid #b8cde3;

            border-radius:6px;

            align-items:center;

            gap:10px;
        }


        .pass-attachment-selected-icon {

            width:32px;

            height:32px;

            display:flex;

            align-items:center;

            justify-content:center;

            background:#dceaf7;

            border-radius:5px;

            flex-shrink:0;

            font-size:16px;
        }


        .pass-attachment-selected-details {

            flex:1;

            min-width:0;
        }


        .pass-attachment-selected-name {

            font-size:13px;

            font-weight:600;

            color:#323130;

            overflow:hidden;

            text-overflow:ellipsis;

            white-space:nowrap;
        }


        .pass-attachment-selected-size {

            margin-top:2px;

            font-size:11px;

            color:#605e5c;
        }


        .pass-attachment-remove {

            width:30px;

            height:30px;

            border:none;

            background:transparent;

            color:#6b7280;

            font-size:20px;

            line-height:1;

            cursor:pointer;

            border-radius:4px;

            flex-shrink:0;
        }


        .pass-attachment-remove:hover {

            background:#e5e7eb;

            color:#111827;
        }


        .pass-attachment-error {

            display:none;

            margin-top:6px;

            color:#a4262c;

            font-size:12px;

            line-height:1.4;
        }


        .pass-attachment-uploading {

            display:none;

            margin-top:8px;

            padding:9px 11px;

            background:#f3f7fb;

            border-left:
                4px solid #0078d4;

            border-radius:4px;

            color:#005a9e;

            font-size:12px;

            font-weight:600;
        }


        .pass-form-actions {

            display:flex;

            justify-content:flex-end;

            gap:10px;

            margin-top:25px;

            padding-top:20px;

            border-top:
                1px solid #edebe9;
        }


        .pass-form-button {

            border:none;

            border-radius:5px;

            padding:11px 20px;

            font-family:
                "Segoe UI",
                Arial,
                sans-serif;

            font-size:14px;

            font-weight:600;

            cursor:pointer;
        }


        .pass-form-cancel {

            background:#f3f2f1;

            color:#323130;
        }


        .pass-form-submit {

            background:#0078d4;

            color:white;
        }


        .pass-form-submit:hover {

            background:#106ebe;
        }


        .pass-form-submit:disabled {

            background:#c8c6c4;

            cursor:not-allowed;
        }


        .pass-form-message {

            display:none;

            padding:
                14px 16px;

            border-radius:6px;

            margin-bottom:20px;

            font-size:13px;
        }


        .pass-form-message-success {

            display:block;

            background:#dff6dd;

            color:#107c10;

            border-left:
                4px solid #107c10;
        }


        .pass-form-message-error {

            display:block;

            background:#fde7e9;

            color:#a4262c;

            border-left:
                4px solid #a4262c;
        }


        .pass-site-loading {

            color:#605e5c;

            font-size:12px;

            padding:
                4px 0;
        }


        .pass-site-error {

            color:#a4262c;

            font-size:12px;

            padding:
                4px 0;
        }


        @media (max-width:900px) {

            .summary-grid {

                grid-template-columns:
                    repeat(2,1fr);
            }


            .controls {

                flex-wrap:wrap;
            }


            .search-box {

                flex-basis:100%;
            }


            .table-container {

                overflow-x:auto;
            }


            .pass-table {

                min-width:800px;
            }


            .detail-grid {

                grid-template-columns:1fr;
            }


            .detail-section-full {

                grid-column:auto;
            }

        }


        @media (max-width:700px) {

            .pass-choice-grid {

                grid-template-columns:1fr;
            }


            .pass-form-grid {

                grid-template-columns:1fr;
            }


            .pass-form-field-full {

                grid-column:auto;
            }


            .pass-home-header {

                padding:28px;
            }


            .pass-form-container {

                padding:20px;
            }

        }


        @media (max-width:550px) {

            .summary-grid {

                grid-template-columns:1fr;
            }


            .dashboard-header h1 {

                font-size:22px;
            }

        }

    `;


    document.head.appendChild(style);


    /*
    ============================================================
    APPLICATION NAVIGATION
    ============================================================
    */

    function renderNavigation() {

        var adminButton = "";


        if (isAdmin) {

            adminButton = `

                <button
                    type="button"
                    class="pass-nav-button pass-nav-button-admin"
                    id="passAdminDashboardButton"
                >
                    Admin Dashboard
                </button>

            `;

        }


        return `

            <div class="pass-app-nav">

                <div class="pass-app-nav-title">
                    Building Pass Request Solution
                </div>

                <div class="pass-app-nav-buttons">

                    <button
                        type="button"
                        class="pass-nav-button"
                        id="passHomeButton"
                    >
                        Home
                    </button>


                    <button
                        type="button"
                        class="pass-nav-button"
                        id="passSubmitButton"
                    >
                        Submit Request
                    </button>


                    <button
                        type="button"
                        class="pass-nav-button pass-nav-button-primary"
                        id="passDashboardButton"
                    >
                        Reporting Dashboard
                    </button>


                    ${adminButton}

                </div>

            </div>

        `;

    }


    /*
    ============================================================
    HOME PAGE
    ============================================================
    */

    function renderHome() {

        dashboard.innerHTML = `

            <div class="pass-home">

                ${renderNavigation()}


                <div class="pass-home-header">

                    <h1>
                        Building Pass Request Solution
                    </h1>

                    <p>
                        Submit a new Building Pass Request or
                        view the reporting dashboard.
                    </p>

                </div>


                <div class="pass-choice-grid">


                    <div
                        class="pass-choice-card"
                        id="openSubmissionCard"
                    >

                        <div class="pass-choice-icon">
                            +
                        </div>

                        <h2>
                            Submit a Pass Request
                        </h2>

                        <p>
                            Complete the Building Pass Request
                            form and submit it directly to the
                            Building Pass register.
                        </p>

                    </div>


                    <div
                        class="pass-choice-card"
                        id="openDashboardCard"
                    >

                        <div class="pass-choice-icon">
                            ☷
                        </div>

                        <h2>
                            Reporting Dashboard
                        </h2>

                        <p>
                            View your current pass requests,
                            statuses, expiry information,
                            sites and access permissions.
                        </p>

                    </div>


                    ${
                        isAdmin
                        ?
                        `
                            <div
                                class="pass-choice-card"
                                id="openAdminDashboardCard"
                            >

                                <div class="pass-choice-icon">
                                    ⚙
                                </div>

                                <h2>
                                    Admin Dashboard
                                </h2>

                                <p>
                                    View and search all access
                                    pass requests across the
                                    Building Pass register.
                                </p>

                            </div>
                        `
                        :
                        ""
                    }


                </div>

            </div>

        `;


        document
            .getElementById(
                "openSubmissionCard"
            )
            .addEventListener(
                "click",
                function () {

                    currentView =
                        "submit";

                    renderSubmissionForm();

                }
            );


        document
            .getElementById(
                "openDashboardCard"
            )
            .addEventListener(
                "click",
                function () {

                    currentView =
                        "dashboard";

                    prepareUserDashboard();

                }
            );


        var adminCard =
            document.getElementById(
                "openAdminDashboardCard"
            );


        if (adminCard) {

            adminCard.addEventListener(
                "click",
                function () {

                    if (!isAdmin) {

                        return;

                    }


                    currentView =
                        "admin";

                    prepareAdminDashboard();

                }
            );

        }


        bindNavigation();

    }


    /*
    ============================================================
    SUBMISSION FORM
    ============================================================
    */

    function renderSubmissionForm() {

        dashboard.innerHTML = `

            <div>

                ${renderNavigation()}

                <div class="dashboard-header">

                    <h1>
                        Building Pass Request Form
                    </h1>

                </div>


                <div class="pass-form-container">

                    <div class="pass-form-header">

                        <h1>
                            Submit Building Pass Request
                        </h1>

                        <p>
                            Please provide the details required
                            for the Building Pass Request. Please note, Building Pass Requests can be submitted for yourself, but will be subject to an approval by your Line Manager, or alternatively your Line Manager can submit on your behalf.
                        </p>

                    </div>


                    <div
                        id="passFormMessage"
                        class="pass-form-message"
                    ></div>


                    <form
                        id="passSubmissionForm"
                        novalidate
                    >


<div class="pass-form-grid">
    <div class="pass-form-field">
        <label class="pass-form-label">
            Are you a new starter?
            <span class="pass-form-required">*</span>
        </label>

        <select id="formNewStarter" class="pass-form-input" required>
            <option value="">Please select</option>
            <option value="yes">Yes</option>
            <option value="no">No</option>
        </select>
    </div>
    
                            <div class="pass-form-field">

                                <label class="pass-form-label">

                                    Pass Requester

                                </label>

                                <input
                                    type="text"
                                    id="formPassRequester"
                                    class="pass-form-input"
                                    readonly
                                >

                                <div class="pass-form-help">

                                    The request will be associated
                                    with your current SharePoint account.

                                </div>

                            </div>
    </div>
    <br>
    <div class="pass-form-grid">
                            <div class="pass-form-field">

                                <label class="pass-form-label">

                                    Pass Holder
                                    <span class="pass-form-required">*</span>

                                </label>


                                <div
                                    class="pass-user-picker"
                                    id="passHolderPicker"
                                >

                                    <input
                                        type="text"
                                        id="formPassHolderSearch"
                                        class="pass-form-input"
                                        placeholder="Start typing a name..."
                                        autocomplete="off"
                                    >


                                    <div
                                        id="passHolderResults"
                                        class="pass-user-results"
                                    ></div>


                                    <div
                                        id="passHolderSelected"
                                        class="pass-user-selected"
                                    ></div>


                                    <input
                                        type="hidden"
                                        id="formPassHolder"
                                        value=""
                                    >

                                </div>
                                





                                <div class="pass-form-help">

                                    Start typing the pass holder's name and select
                                    the correct person from the results.

                                </div>

                            </div>

                            <div class="pass-form-field">

                                <label class="pass-form-label">

                                    Site
                                    <span class="pass-form-required">*</span>

                                </label>

                                <select
                                    id="formSite"
                                    class="pass-form-select"
                                    required
                                >

                                    <option value="">
                                        Loading sites...
                                    </option>

                                </select>

                                <div
                                    id="passSiteStatus"
                                    class="pass-site-loading"
                                >
                                    Loading available sites...
                                </div>

                            </div>


                            <div class="pass-form-field">

                                <label class="pass-form-label">

                                    Clearance Level
                                    <span class="pass-form-required">*</span>

                                </label>

                                <select
                                    id="formClearance"
                                    class="pass-form-select"
                                    required
                                >

                                    <option value="">
                                        Select clearance
                                    </option>

                                    <option value="BPSS">
                                        BPSS (Baseline Personnel Security Standard) - Blue stripe
                                    </option>

                                    <option value="CTC">
                                        CTC (Counter Terrorism Check) - Red stripe
                                    </option>

                                    <option value="SC">
                                        SC (Security Check) - Yellow stripe
                                    </option>

                                    <option value="DV">
                                        DV (Developed Vetting) - Green stripe
                                    </option>

                                </select>

                            </div>


<div
    class="pass-form-field"
    id="nationalInsuranceField"
    style="display: none;"
>
    <label class="pass-form-label">
        National Insurance Number
        <span class="pass-form-required">*</span>
    </label>

    <input
        type="text"
        id="formNationalInsurance"
        class="pass-form-input"
    >
</div>

</div>
<br>

                            <!--
                            ====================================================
                            PDF ATTACHMENT
                            ====================================================
                            -->
<div class="pass-form-grid">

                            <div
                                class="pass-form-field pass-form-field-full" id="formAttachmentFull"
                            >

                                <label class="pass-form-label">

                                    Optional supporting document (PDF)

                                </label>


                                <div class="pass-attachment-wrapper">

                                    <input
                                        type="file"
                                        id="formAttachmentPdf"
                                        class="pass-attachment-input"
                                        accept=".pdf,application/pdf"
                                    >


                                    <div
                                        class="pass-form-help"
                                    >

                                        Optional. PDF only.
                                        Maximum file size: 10 MB.

                                    </div>


                                    <div
                                        id="passAttachmentPdfError"
                                        class="pass-attachment-error"
                                    ></div>


                                    <div
                                        id="passAttachmentPdfSelected"
                                        class="pass-attachment-selected"
                                    ></div>


                                    <div
                                        id="passAttachmentPdfUploading"
                                        class="pass-attachment-uploading"
                                    >
                                        Uploading attachment...

                                    </div>

                                </div>

                            </div>
</div>
<br>
<div class="pass-form-grid">
                            <div class="pass-form-field">

                                <label class="pass-form-label">

                                    Required Date
                                    <span class="pass-form-required">*</span>

                                </label>

                                <input
                                    type="date"
                                    id="formRequiredDate"
                                    class="pass-form-input"
                                    required
                                >

                            </div>


                            <div class="pass-form-field">

                                <label class="pass-form-label">

                                    Expiry Date

                                </label>

                                <input
                                    type="date"
                                    id="formExpiryDate"
                                    class="pass-form-input"
                                >

                            </div>




                            <div class="pass-form-field">

                                <label class="pass-form-label">

                                    Access Permissions
                                    <span class="pass-form-required">*</span>

                                </label>

                                <select
                                    id="formAccess"
                                    class="pass-form-select pass-access-select"
                                    multiple
                                    required
                                    disabled
                                >

                                    <option value="">
                                        Select a site first
                                    </option>

                                </select>

                                <div
                                    id="passAccessStatus"
                                    class="pass-form-help"
                                >
                                    Select a site to see available access areas.

                                </div>

                            </div>


                            <div
                                class="pass-form-field pass-form-field-full"
                            >

                                <label class="pass-form-label">

                                    Justification
                                    <span class="pass-form-required">*</span>

                                </label>

                                <textarea
                                    id="formJustification"
                                    class="pass-form-textarea"
                                    required
                                ></textarea>

                            </div>


                            <!--
                            ====================================================
                            ATTACHMENT
                            ====================================================
                            -->

                            <div
                                class="pass-form-field pass-form-field-full"
                            >

                                <label class="pass-form-label">

                                    Optional picture of passholder

                                </label>


                                <div class="pass-attachment-wrapper">

                                    <input
                                        type="file"
                                        id="formAttachment"
                                        class="pass-attachment-input"
                                        accept=".jpg,.jpeg,.png,image/jpeg,image/png"
                                    >


                                    <div
                                        class="pass-form-help"
                                    >

                                        Optional. JPG, JPEG or PNG only.
                                        Maximum file size: 10 MB.

                                    </div>


                                    <div
                                        id="passAttachmentError"
                                        class="pass-attachment-error"
                                    ></div>


                                    <div
                                        id="passAttachmentSelected"
                                        class="pass-attachment-selected"
                                    ></div>


                                    <div
                                        id="passAttachmentUploading"
                                        class="pass-attachment-uploading"
                                    >
                                        Uploading attachment...

                                    </div>

                                </div>

                            </div>





                        </div>


                        <div class="pass-form-actions">

                            <button
                                type="button"
                                id="passFormCancel"
                                class="pass-form-button pass-form-cancel"
                            >
                                Cancel
                            </button>


                            <button
                                type="submit"
                                id="passFormSubmit"
                                class="pass-form-button pass-form-submit"
                            >
                                Submit Request
                            </button>

                        </div>


                    </form>

                </div>

            </div>

        `;


        bindNavigation();


        var requesterField =
            document.getElementById(
                "formPassRequester"
            );


        if (requesterField) {

            requesterField.value =
                currentUserEmail ||
                currentUserDisplayName ||
                "";

        }


        var siteSelect =
            document.getElementById(
                "formSite"
            );


        var siteStatus =
            document.getElementById(
                "passSiteStatus"
            );


        var accessSelect =
            document.getElementById(
                "formAccess"
            );


const newStarterSelect = document.getElementById("formNewStarter");
const clearanceSelect = document.getElementById("formClearance");
const nationalInsuranceField = document.getElementById("nationalInsuranceField");
const nationalInsuranceInput = document.getElementById("formNationalInsurance");
const attachmentPdfField = document.getElementById("formAttachmentPdf");
const attachmentPdfField1 = document.getElementById("formAttachmentFull");

function updateNationalInsuranceVisibility() {
    const isNewStarter = newStarterSelect.value === "yes";
    const clearance = clearanceSelect.value;

    const clearanceRequiresNI = ["CTC", "SC", "DV"].includes(clearance);
    const shouldShowNI = isNewStarter && clearanceRequiresNI;

    // Show/hide National Insurance field
    if (shouldShowNI) {
        nationalInsuranceField.style.display = "";
        nationalInsuranceInput.required = true;
    } else {
        nationalInsuranceField.style.display = "none";
        nationalInsuranceInput.required = false;
        nationalInsuranceInput.value = "";
    }

    // Show/hide Attachment PDF field
    if (shouldShowNI) {
        attachmentPdfField1.style.display = "";
    } else {
        attachmentPdfField1.style.display = "none";
    }
}

newStarterSelect.addEventListener("change", updateNationalInsuranceVisibility);
clearanceSelect.addEventListener("change", updateNationalInsuranceVisibility);

// Set the correct initial state
updateNationalInsuranceVisibility();



        /*
        --------------------------------------------------------
        PASS HOLDER USER SEARCH
        --------------------------------------------------------
        */

        selectedPassHolder =
            null;


        var passHolderSearch =
            document.getElementById(
                "formPassHolderSearch"
            );


        if (passHolderSearch) {

            passHolderSearch.addEventListener(
                "input",
                function () {

                    var searchText =
                        this.value.trim();


                    if (
                        selectedPassHolder
                    ) {

                        selectedPassHolder =
                            null;


                        var hiddenField =
                            document.getElementById(
                                "formPassHolder"
                            );


                        if (hiddenField) {

                            hiddenField.value =
                                "";

                        }

                    }


                    clearTimeout(
                        passHolderSearchTimer
                    );


                    passHolderSearchTimer =
                        setTimeout(
                            function () {

                                searchPassHolders(
                                    searchText
                                );

                            },
                            300
                        );

                }
            );

        }


        /*
        --------------------------------------------------------
        SITE CHANGE -> UPDATE ACCESS OPTIONS
        --------------------------------------------------------
        */

        if (siteSelect) {

            siteSelect.addEventListener(
                "change",
                function () {

                    updateAccessPermissionsDropdown(
                        true
                    );

                }
            );

        }


        populateSiteDropdown();


        /*
        --------------------------------------------------------
        INITIAL ACCESS DROPDOWN STATE
        --------------------------------------------------------
        */

        if (accessSelect) {

            updateAccessPermissionsDropdown(
                true
            );

        }


        /*
        --------------------------------------------------------
        ATTACHMENT FILE INPUT
        --------------------------------------------------------
        */

        var attachmentInput =
            document.getElementById(
                "formAttachment"
            );


        if (attachmentInput) {

            attachmentInput.addEventListener(
                "change",
                function () {

                    validateAndDisplayAttachment(
                        this
                    );

                }
            );

        }


        /*
        --------------------------------------------------------
        PDF ATTACHMENT FILE INPUT
        --------------------------------------------------------
        */

        var attachmentPdfInput =
            document.getElementById(
                "formAttachmentPdf"
            );


        if (attachmentPdfInput) {

            attachmentPdfInput.addEventListener(
                "change",
                function () {

                    validateAndDisplayPdfAttachment(
                        this
                    );

                }
            );

        }


        document
            .getElementById(
                "passFormCancel"
            )
            .addEventListener(
                "click",
                function () {

                    currentView =
                        "home";

                    renderHome();

                }
            );


        document
            .getElementById(
                "passSubmissionForm"
            )
            .addEventListener(
                "submit",
                function (event) {

                    event.preventDefault();

                    submitPassRequest();

                }
            );

    }


    /*
    ============================================================
    ATTACHMENT HELPERS
    ============================================================
    */


function validateAndDisplayAttachment(input) {

    var selectedContainer =
        document.getElementById(
            "passAttachmentSelected"
        );

    showAttachmentError("");

    if (
        !input ||
        !input.files ||
        input.files.length === 0
    ) {

        if (selectedContainer) {

            selectedContainer.innerHTML = "";
            selectedContainer.style.display = "none";

        }

        return [];

    }


    /*
    ------------------------------------------------------------
    Maximum 2 files
    ------------------------------------------------------------
    */

    if (input.files.length > 2) {

        input.value = "";

        if (selectedContainer) {

            selectedContainer.innerHTML = "";
            selectedContainer.style.display = "none";

        }

        showAttachmentError(
            "You can upload a maximum of 2 files."
        );

        return [];

    }


    var files = Array.from(input.files);


    /*
    ------------------------------------------------------------
    Validate every selected file
    ------------------------------------------------------------
    */

    for (var i = 0; i < files.length; i++) {

        var validation =
            validateAttachmentFile(
                files[i]
            );


        if (!validation.valid) {

            input.value = "";

            if (selectedContainer) {

                selectedContainer.innerHTML = "";
                selectedContainer.style.display = "none";

            }

            showAttachmentError(
                files[i].name +
                ": " +
                validation.message
            );

            return [];

        }

    }


    /*
    ------------------------------------------------------------
    Display selected files
    ------------------------------------------------------------
    */

    if (selectedContainer) {

        selectedContainer.innerHTML = "";


        files.forEach(
            function (file, index) {

                var fileElement =
                    document.createElement("div");

                fileElement.className =
                    "pass-attachment-selected";


                fileElement.innerHTML = `

                    <div class="pass-attachment-selected-icon">
                        📎
                    </div>

                    <div class="pass-attachment-selected-details">

                        <div class="pass-attachment-selected-name">
                            ${escapeHtml(file.name)}
                        </div>

                        <div class="pass-attachment-selected-size">
                            ${formatFileSize(file.size)}
                        </div>

                    </div>

                `;


                selectedContainer.appendChild(
                    fileElement
                );

            }
        );


        /*
        --------------------------------------------------------
        Add one remove button for all selected files
        --------------------------------------------------------
        */

        var removeButton =
            document.createElement("button");

        removeButton.type = "button";
        removeButton.className =
            "pass-attachment-remove";
        removeButton.title =
            "Remove attachments";
        removeButton.textContent = "×";


        removeButton.addEventListener(
            "click",
            function () {

                clearAttachment();

            }
        );


        selectedContainer.appendChild(
            removeButton
        );


        selectedContainer.style.display =
            "flex";

    }


    return files;

}


    function clearAttachment() {

        var input =
            document.getElementById(
                "formAttachment"
            );


        var selectedContainer =
            document.getElementById(
                "passAttachmentSelected"
            );


        var uploadingContainer =
            document.getElementById(
                "passAttachmentUploading"
            );


        if (input) {

            input.value =
                "";

        }


        if (selectedContainer) {

            selectedContainer.innerHTML =
                "";

            selectedContainer.style.display =
                "none";

        }


        if (uploadingContainer) {

            uploadingContainer.style.display =
                "none";

        }


        showAttachmentError("");

    }


    /*
    ============================================================
    PDF ATTACHMENT HELPERS
    ============================================================

    Mirrors validateAndDisplayAttachment / clearAttachment above,
    but targets the dedicated PDF file input and its own
    error / selected-file / uploading containers so the two
    upload widgets do not interfere with each other.
    ============================================================
    */

    function showPdfAttachmentError(message) {

        var errorContainer =
            document.getElementById(
                "passAttachmentPdfError"
            );


        if (!errorContainer) {

            return;

        }


        errorContainer.textContent =
            message ||
            "";


        errorContainer.style.display =
            message
                ? "block"
                : "none";

    }


    function validateAndDisplayPdfAttachment(input) {

        var selectedContainer =
            document.getElementById(
                "passAttachmentPdfSelected"
            );

        showPdfAttachmentError("");

        if (
            !input ||
            !input.files ||
            input.files.length === 0
        ) {

            if (selectedContainer) {

                selectedContainer.innerHTML = "";
                selectedContainer.style.display = "none";

            }

            return [];

        }


        /*
        ------------------------------------------------------------
        Only one PDF may be attached via this input.
        ------------------------------------------------------------
        */

        if (input.files.length > 1) {

            input.value = "";

            if (selectedContainer) {

                selectedContainer.innerHTML = "";
                selectedContainer.style.display = "none";

            }

            showPdfAttachmentError(
                "You can upload a maximum of 1 PDF file."
            );

            return [];

        }


        var files = Array.from(input.files);


        /*
        ------------------------------------------------------------
        Validate the selected file (size, type, extension).
        ------------------------------------------------------------
        */

        for (var i = 0; i < files.length; i++) {

            var validation =
                validateAttachmentFile(
                    files[i]
                );


            if (!validation.valid) {

                input.value = "";

                if (selectedContainer) {

                    selectedContainer.innerHTML = "";
                    selectedContainer.style.display = "none";

                }

                showPdfAttachmentError(
                    files[i].name +
                    ": " +
                    validation.message
                );

                return [];

            }

        }


        /*
        ------------------------------------------------------------
        Display the selected file
        ------------------------------------------------------------
        */

        if (selectedContainer) {

            selectedContainer.innerHTML = "";


            files.forEach(
                function (file, index) {

                    var fileElement =
                        document.createElement("div");

                    fileElement.className =
                        "pass-attachment-selected";


                    fileElement.innerHTML = `

                        <div class="pass-attachment-selected-icon">
                            📎
                        </div>

                        <div class="pass-attachment-selected-details">

                            <div class="pass-attachment-selected-name">
                                ${escapeHtml(file.name)}
                            </div>

                            <div class="pass-attachment-selected-size">
                                ${formatFileSize(file.size)}
                            </div>

                        </div>

                    `;


                    selectedContainer.appendChild(
                        fileElement
                    );

                }
            );


            var removeButton =
                document.createElement("button");

            removeButton.type = "button";
            removeButton.className =
                "pass-attachment-remove";
            removeButton.title =
                "Remove attachment";
            removeButton.textContent = "×";


            removeButton.addEventListener(
                "click",
                function () {

                    clearPdfAttachment();

                }
            );


            selectedContainer.appendChild(
                removeButton
            );


            selectedContainer.style.display =
                "flex";

        }


        return files;

    }


    function clearPdfAttachment() {

        var input =
            document.getElementById(
                "formAttachmentPdf"
            );


        var selectedContainer =
            document.getElementById(
                "passAttachmentPdfSelected"
            );


        var uploadingContainer =
            document.getElementById(
                "passAttachmentPdfUploading"
            );


        if (input) {

            input.value =
                "";

        }


        if (selectedContainer) {

            selectedContainer.innerHTML =
                "";

            selectedContainer.style.display =
                "none";

        }


        if (uploadingContainer) {

            uploadingContainer.style.display =
                "none";

        }


        showPdfAttachmentError("");

    }


    /*
    ============================================================
    SHAREPOINT ATTACHMENT FILENAME
    ============================================================

    SharePoint attachment URLs use an OData string inside
    FileName='...'.

    Apostrophes therefore need to be escaped.
    ============================================================
    */

    function escapeSharePointAttachmentFileName(
        fileName
    ) {

        var safeName =
            String(
                fileName || ""
            );


        /*
        --------------------------------------------------------
        Remove path information. A browser normally supplies
        only the filename, but this makes the function safer.
        --------------------------------------------------------
        */

        safeName =
            safeName
                .split("\\")
                .pop()
                .split("/")
                .pop();


        /*
        --------------------------------------------------------
        Keep the filename reasonably sized for SharePoint.
        --------------------------------------------------------
        */

        if (
            safeName.length >
            180
        ) {

            var extension =
                getAttachmentExtension(
                    safeName
                );


            var baseName =
                safeName.substring(
                    0,
                    safeName.length -
                    extension.length
                );


            baseName =
                baseName.substring(
                    0,
                    180 -
                    extension.length
                );


            safeName =
                baseName +
                extension;

        }


        /*
        --------------------------------------------------------
        OData escaping for apostrophes.
        --------------------------------------------------------
        */

        safeName =
            safeName.replace(
                /'/g,
                "''"
            );


        return encodeURIComponent(
            safeName
        )
        /*
        --------------------------------------------------------
        encodeURIComponent intentionally leaves some characters
        alone. Encode apostrophes explicitly as well.
        --------------------------------------------------------
        */
        .replace(
            /'/g,
            "%27"
        );

    }


    /*
    ============================================================
    UPLOAD ATTACHMENT TO SHAREPOINT ITEM
    ============================================================
    */

    function uploadAttachmentToItem(
        itemId,
        file,
        digest
    ) {

        if (!file) {

            return Promise.resolve();

        }


        var validation =
            validateAttachmentFile(
                file
            );


        if (!validation.valid) {

            return Promise.reject(
                new Error(
                    validation.message
                )
            );

        }


        var encodedFileName =
            escapeSharePointAttachmentFileName(
                file.name
            );


        var attachmentUrl =
            LIST_API_URL +
            "(" +
            encodeURIComponent(
                itemId
            ) +
            ")/AttachmentFiles/add(FileName='" +
            encodedFileName +
            "')";


        console.log(
            "Uploading SharePoint attachment:",
            file.name,
            "to item:",
            itemId
        );


        return fetch(
            attachmentUrl,
            {
                method: "POST",

                credentials:
                    "same-origin",

                headers: {

                    "Accept":
                        "application/json;odata=verbose",

                    "Content-Type":
                        "application/octet-stream",

                    "X-RequestDigest":
                        digest

                },

                body:
                    file

            }
        )
        .then(
            function (response) {

                if (!response.ok) {

                    return response.text()
                        .then(
                            function (text) {

                                console.error(
                                    "SharePoint attachment upload response:",
                                    text
                                );


                                throw new Error(
                                    "SharePoint returned HTTP " +
                                    response.status +
                                    " while uploading the attachment. " +
                                    text
                                );

                            }
                        );

                }


                return response.json();

            }
        )
        .then(
            function (result) {

                console.log(
                    "Attachment uploaded successfully.",
                    result
                );


                return result;

            }
        );

    }


    /*
    ============================================================
    POPULATE SITE DROPDOWN
    ============================================================
    */

    function populateSiteDropdown() {

        var siteSelect =
            document.getElementById(
                "formSite"
            );


        var siteStatus =
            document.getElementById(
                "passSiteStatus"
            );


        if (!siteSelect) {

            return;

        }


        siteSelect.innerHTML = `

            <option value="">
                Select site
            </option>

        `;


        availableSites.forEach(
            function (site) {

                var option =
                    document.createElement(
                        "option"
                    );


                option.value =
                    site;


                option.textContent =
                    site;


                siteSelect.appendChild(
                    option
                );

            }
        );


        if (siteStatus) {

            siteStatus.className =
                "pass-site-loading";


            siteStatus.textContent =
                availableSites.length +
                " site" +
                (
                    availableSites.length === 1
                        ? ""
                        : "s"
                ) +
                " available.";

        }


        updateAccessPermissionsDropdown(
            true
        );

    }


    /*
    ============================================================
    GET REQUEST DIGEST
    ============================================================
    */

    function getRequestDigest() {

        return fetch(
            SITE_URL +
            "/_api/contextinfo",
            {
                method: "POST",

                headers: {

                    "Accept":
                        "application/json;odata=verbose",

                    "Content-Type":
                        "application/json;odata=verbose"

                },

                credentials:
                    "same-origin"

            }
        )
        .then(
            function (response) {

                if (!response.ok) {

                    throw new Error(
                        "Unable to obtain SharePoint request digest. HTTP " +
                        response.status
                    );

                }


                return response.json();

            }
        )
        .then(
            function (json) {

                return json
                    .d
                    .GetContextWebInformation
                    .FormDigestValue;

            }
        );

    }


    /*
    ============================================================
    SEARCH SHAREPOINT USERS
    ============================================================
    */

    function searchPassHolders(searchText) {

        var resultsContainer =
            document.getElementById(
                "passHolderResults"
            );


        if (!resultsContainer) {

            return;

        }


        searchText =
            searchText.trim();


        if (searchText.length < 2) {

            resultsContainer.innerHTML =
                "";

            resultsContainer.style.display =
                "none";

            return;

        }


        resultsContainer.innerHTML = `

            <div class="pass-user-no-results">

                Searching...

            </div>

        `;


        resultsContainer.style.display =
            "block";


        var queryParams = {

            __metadata: {

                type:
                    "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters"

            },

            AllowEmailAddresses: true,

            AllowMultipleEntities: false,

            AllUrlZones: false,

            MaximumEntitySuggestions: 10,

            PrincipalSource: 15,

            PrincipalType: 1,

            QueryString: searchText

        };


        getRequestDigest()

        .then(
            function (digest) {

                return fetch(
                    SITE_URL +
                    "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser",
                    {
                        method: "POST",

                        credentials:
                            "same-origin",

                        headers: {

                            "Accept":
                                "application/json;odata=verbose",

                            "Content-Type":
                                "application/json;odata=verbose",

                            "X-RequestDigest":
                                digest

                        },

                        body:
                            JSON.stringify(
                                {
                                    queryParams:
                                        queryParams
                                }
                            )

                    }
                );

            }
        )
        .then(
            function (response) {

                console.log(
                    "People picker search HTTP status:",
                    response.status
                );


                if (!response.ok) {

                    return response.text()
                        .then(
                            function (text) {

                                console.error(
                                    "People picker search response:",
                                    text
                                );


                                throw new Error(
                                    "People picker search returned HTTP " +
                                    response.status +
                                    ". " +
                                    text
                                );

                            }
                        );

                }


                return response.json();

            }
        )
        .then(
            function (json) {

                var raw =
                    json &&
                    json.d &&
                    json.d.ClientPeoplePickerSearchUser;


                var entities = [];


                try {

                    entities =
                        JSON.parse(
                            raw
                        ) ||
                        [];

                }
                catch (error) {

                    console.error(
                        "Unable to parse people picker response.",
                        error,
                        raw
                    );

                    entities = [];

                }


                console.log(
                    "People picker results:",
                    entities
                );


                renderPassHolderResults(
                    entities
                );

            }
        )
        .catch(
            function (error) {

                console.error(
                    "Pass holder search failed.",
                    error
                );


                resultsContainer.innerHTML = `

                    <div class="pass-user-search-error">

                        Unable to search for users.

                    </div>

                `;


                resultsContainer.style.display =
                    "block";

            }
        );

    }


    /*
    ============================================================
    RENDER PASS HOLDER SEARCH RESULTS
    ============================================================
    */

    function renderPassHolderResults(entities) {

        var resultsContainer =
            document.getElementById(
                "passHolderResults"
            );


        if (!resultsContainer) {

            return;

        }


        var users = (entities || [])

            .map(
                function (entity) {

                    var entityData =
                        entity.EntityData ||
                        {};


                    var email =
                        entityData.Email ||
                        "";


                    if (
                        !email &&
                        entity.Key &&
                        entity.Key.indexOf("@") !== -1
                    ) {

                        email =
                            entity.Key;

                    }


                    return {

                        displayName:
                            entity.DisplayText ||
                            entity.Description ||
                            "",

                        email:
                            email,

                        accountName:
                            entity.Key ||
                            ""

                    };

                }
            )

            .filter(
                function (user) {

                    return !!user.displayName;

                }
            );


        if (users.length === 0) {

            resultsContainer.innerHTML = `

                <div class="pass-user-no-results">

                    No matching users found.

                </div>

            `;


            resultsContainer.style.display =
                "block";


            return;

        }


        resultsContainer.innerHTML =
            users
                .map(
                    function (user, index) {

                        return `

                            <div
                                class="pass-user-result"
                                data-user-index="${index}"
                            >

                                <div class="pass-user-result-icon">

                                    👤

                                </div>


                                <div class="pass-user-result-details">

                                    <div class="pass-user-result-name">

                                        ${escapeHtml(
                                            user.displayName
                                        )}

                                    </div>


                                    ${
                                        user.email

                                        ?

                                        `

                                            <div class="pass-user-result-email">

                                                ${escapeHtml(
                                                    user.email
                                                )}

                                            </div>

                                        `

                                        :

                                        ""

                                    }

                                </div>

                            </div>

                        `;

                    }
                )
                .join("");


        resultsContainer.style.display =
            "block";


        var resultElements =
            resultsContainer.querySelectorAll(
                ".pass-user-result"
            );


        resultElements.forEach(
            function (element) {

                element.addEventListener(
                    "click",
                    function () {

                        var index =
                            parseInt(
                                this.getAttribute(
                                    "data-user-index"
                                ),
                                10
                            );


                        selectPassHolder(
                            users[index]
                        );

                    }
                );

            }
        );

    }


    /*
    ============================================================
    SELECT PASS HOLDER
    ============================================================
    */

    function selectPassHolder(user) {

        selectedPassHolder =
            user;


        var hiddenField =
            document.getElementById(
                "formPassHolder"
            );


        var searchInput =
            document.getElementById(
                "formPassHolderSearch"
            );


        var resultsContainer =
            document.getElementById(
                "passHolderResults"
            );


        var selectedContainer =
            document.getElementById(
                "passHolderSelected"
            );


        if (hiddenField) {

            hiddenField.value =
                user.email;

        }


        if (searchInput) {

            searchInput.value =
                "";

            searchInput.style.display =
                "none";

        }


        if (resultsContainer) {

            resultsContainer.innerHTML =
                "";

            resultsContainer.style.display =
                "none";

        }


        if (selectedContainer) {

            selectedContainer.innerHTML = `

                <div class="pass-selected-user">

                    <div class="pass-selected-user-icon">

                        👤

                    </div>


                    <div class="pass-selected-user-details">

                        <div class="pass-selected-user-name">

                            ${escapeHtml(
                                user.displayName
                            )}

                        </div>


                        ${
                            user.email

                            ?

                            `

                                <div class="pass-selected-user-email">

                                    ${escapeHtml(
                                        user.email
                                    )}

                                </div>

                            `

                            :

                            ""

                        }

                    </div>


                    <button
                        type="button"
                        class="pass-selected-user-remove"
                        id="removePassHolder"
                        title="Remove selected user"
                    >

                        ×

                    </button>

                </div>

            `;


            selectedContainer.style.display =
                "block";


            document
                .getElementById(
                    "removePassHolder"
                )
                .addEventListener(
                    "click",
                    function () {

                        clearPassHolder();

                    }
                );

        }

    }


    /*
    ============================================================
    CLEAR SELECTED PASS HOLDER
    ============================================================
    */

    function clearPassHolder() {

        selectedPassHolder =
            null;


        var searchInput =
            document.getElementById(
                "formPassHolderSearch"
            );


        var resultsContainer =
            document.getElementById(
                "passHolderResults"
            );


        var selectedContainer =
            document.getElementById(
                "passHolderSelected"
            );


        var hiddenField =
            document.getElementById(
                "formPassHolder"
            );


        if (hiddenField) {

            hiddenField.value =
                "";

        }


        if (selectedContainer) {

            selectedContainer.innerHTML =
                "";

            selectedContainer.style.display =
                "none";

        }


        if (resultsContainer) {

            resultsContainer.innerHTML =
                "";

            resultsContainer.style.display =
                "none";

        }


        if (searchInput) {

            searchInput.value =
                "";

            searchInput.style.display =
                "block";

            searchInput.focus();

        }

    }


    /*
    ============================================================
    SUBMIT PASS REQUEST
    ============================================================
    */

    function submitPassRequest() {

        var submitButton =
            document.getElementById(
                "passFormSubmit"
            );


        var message =
            document.getElementById(
                "passFormMessage"
            );


        var passHolder =
            document.getElementById(
                "formPassHolder"
            )
            .value;


        var NI =
            document.getElementById(
                "formNationalInsurance"
            )
            .value;


        var site =
            document.getElementById(
                "formSite"
            )
            .value;


        var clearance =
            document.getElementById(
                "formClearance"
            )
            .value;


        var requiredDate =
            document.getElementById(
                "formRequiredDate"
            )
            .value;


        var expiryDate =
            document.getElementById(
                "formExpiryDate"
            )
            .value;


        var requester =
            document.getElementById(
                "formPassRequester"
            )
            .value
            .trim();


        var selectedAccess =
            getSelectedAccessPermissions();


        var access =
            selectedAccess.join(", ");


        var justification =
            document.getElementById(
                "formJustification"
            )
            .value
            .trim();


/*
--------------------------------------------------------
GET ATTACHMENTS
--------------------------------------------------------
*/

var attachmentInput =
    document.getElementById(
        "formAttachment"
    );


var attachmentPdfInput =
    document.getElementById(
        "formAttachmentPdf"
    );


var attachmentFiles =
    [];


if (
    attachmentInput &&
    attachmentInput.files &&
    attachmentInput.files.length > 0
) {

    attachmentFiles =
        attachmentFiles.concat(
            Array.from(
                attachmentInput.files
            )
        );

}


/*
--------------------------------------------------------
Add any file selected via the dedicated PDF input so
that BOTH the image and the PDF are uploaded to the
same SharePoint item's Attachments column.
--------------------------------------------------------
*/

if (
    attachmentPdfInput &&
    attachmentPdfInput.files &&
    attachmentPdfInput.files.length > 0
) {

    attachmentFiles =
        attachmentFiles.concat(
            Array.from(
                attachmentPdfInput.files
            )
        );

}


/*
--------------------------------------------------------
VALIDATE ATTACHMENTS
--------------------------------------------------------
*/

/*
--------------------------------------------------------
Maximum 2 files (1 image + 1 PDF)
--------------------------------------------------------
*/

if (
    attachmentFiles.length > 2
) {

    message.className =
        "pass-form-message " +
        "pass-form-message-error";

    message.innerHTML =
        "You can upload a maximum of 2 files.";

    return;

}

/*
--------------------------------------------------------
Validate every attachment
--------------------------------------------------------
*/

for (
    var i = 0;
    i < attachmentFiles.length;
    i++
) {

    var attachmentValidation =
        validateAttachmentFile(
            attachmentFiles[i]
        );


if (
    !attachmentValidation.valid
) {

    message.className =
        "pass-form-message " +
        "pass-form-message-error";

    message.innerHTML =
        escapeHtml(
            attachmentFiles[i].name +
            ": " +
            attachmentValidation.message
        );

    return;

}

}


        /*
        --------------------------------------------------------
        VALIDATE REQUIRED FORM FIELDS
        --------------------------------------------------------
        */

        if (
            !selectedPassHolder ||
            !passHolder ||
            !site ||
            !clearance ||
            !requiredDate ||
            selectedAccess.length === 0 ||
            !justification
        ) {

            message.className =
                "pass-form-message " +
                "pass-form-message-error";


            message.innerHTML =
                "Please complete all required fields and select at least one Access Permission.";


            return;

        }


        submitButton.disabled =
            true;


        submitButton.textContent =
            "Submitting...";


        message.className =
            "pass-form-message";


        message.innerHTML =
            "";


        var uploadingContainer =
            document.getElementById(
                "passAttachmentUploading"
            );


        var uploadingPdfContainer =
            document.getElementById(
                "passAttachmentPdfUploading"
            );


        if (uploadingContainer) {

            uploadingContainer.style.display =
                "none";

        }


        if (uploadingPdfContainer) {

            uploadingPdfContainer.style.display =
                "none";

        }


        var item = {

            "__metadata": {

                "type":
                    LIST_ITEM_TYPE

            },

            Pass_Holder:
                passHolder,

            NI:
                NI,

            Site:
                site,

            Clearance_Level:
                clearance,

            Required_Date:
                requiredDate,

            Access:
                access,

            Justification:
                justification

        };


        if (expiryDate) {

            item.Expiry_Date =
                expiryDate;

        }


        if (requester) {

            item.Pass_Requester =
                requester;

        }


        /*
        --------------------------------------------------------
        GET REQUEST DIGEST ONCE.
        --------------------------------------------------------

        The same digest is then used for:

        1. Creating the SharePoint list item.
        2. Uploading the attachment.

        --------------------------------------------------------
        */

        getRequestDigest()

        .then(
            function (digest) {

                return fetch(
                    LIST_API_URL,
                    {
                        method: "POST",

                        credentials:
                            "same-origin",

                        headers: {

                            "Accept":
                                "application/json;odata=verbose",

                            "Content-Type":
                                "application/json;odata=verbose",

                            "X-RequestDigest":
                                digest

                        },

                        body:
                            JSON.stringify(
                                item
                            )

                    }
                )
                .then(
                    function (response) {

                        if (!response.ok) {

                            return response.text()
                                .then(
                                    function (text) {

                                        throw new Error(
                                            "SharePoint returned HTTP " +
                                            response.status +
                                            ". " +
                                            text
                                        );

                                    }
                                );

                        }


                        return response.json();

                    }
                )
                .then(
                    function (result) {

                        /*
                        ------------------------------------------------
                        SHAREPOINT RETURNS THE NEW ITEM IN result.d
                        ------------------------------------------------
                        */

                        var createdItem =
                            result &&
                            result.d
                                ? result.d
                                : null;


                        var itemId =
                            createdItem &&
                            createdItem.ID
                                ? createdItem.ID
                                : null;


                        if (!itemId) {

                            throw new Error(
                                "The request was created, but SharePoint did not return the new item ID. The attachment could not be uploaded."
                            );

                        }


                        console.log(
                            "Pass request successfully created. Item ID:",
                            itemId
                        );

/*
------------------------------------------------
IF THERE ARE NO ATTACHMENTS, WE ARE FINISHED.
------------------------------------------------
*/

if (
    attachmentFiles.length === 0
) {

    return {

        itemId:
            itemId,

        attachmentUploaded:
            false,

        attachmentCount:
            0

    };

}


/*
------------------------------------------------
SHOW UPLOAD STATUS
------------------------------------------------
*/

if (uploadingContainer) {

    uploadingContainer.style.display =
        "block";

    uploadingContainer.textContent =
        "Uploading attachments...";

}


if (uploadingPdfContainer) {

    uploadingPdfContainer.style.display =
        "block";

    uploadingPdfContainer.textContent =
        "Uploading attachments...";

}


submitButton.textContent =
    "Uploading attachments...";


/*
------------------------------------------------
UPLOAD ALL FILES TO THE NEW ITEM'S ATTACHMENTS
------------------------------------------------
*/

return attachmentFiles
    .reduce(
        function (promise, file, index) {

            return promise.then(
                function () {

                    if (uploadingContainer) {

                        uploadingContainer.textContent =
                            "Uploading attachment " +
                            (index + 1) +
                            " of " +
                            attachmentFiles.length +
                            "...";

                    }


                    if (uploadingPdfContainer) {

                        uploadingPdfContainer.textContent =
                            "Uploading attachment " +
                            (index + 1) +
                            " of " +
                            attachmentFiles.length +
                            "...";

                    }


                    return uploadAttachmentToItem(
                        itemId,
                        file,
                        digest
                    );

                }
            );

        },
        Promise.resolve()
    )
    .then(
        function () {

            if (uploadingContainer) {

                uploadingContainer.textContent =
                    attachmentFiles.length +
                    " attachment" +
                    (
                        attachmentFiles.length === 1
                            ? ""
                            : "s"
                    ) +
                    " uploaded successfully.";

            }


            if (uploadingPdfContainer) {

                uploadingPdfContainer.textContent =
                    attachmentFiles.length +
                    " attachment" +
                    (
                        attachmentFiles.length === 1
                            ? ""
                            : "s"
                    ) +
                    " uploaded successfully.";

            }


            return {

                itemId:
                    itemId,

                attachmentUploaded:
                    true,

                attachmentCount:
                    attachmentFiles.length

            };

        }
    )
    .catch(
        function (attachmentError) {

            /*
            ------------------------------------------------
            IMPORTANT:
            The list item has already been created.

            Therefore we report the attachment error
            separately rather than pretending that the
            entire request failed.
            ------------------------------------------------
            */

            throw new Error(
                "The pass request was created successfully as SharePoint item #" +
                itemId +
                ", but one or more attachments could not be uploaded. " +
                attachmentError.message
            );

        }
    );

                    }
                );

            }
        )

        .then(
            function (result) {

                console.log(
                    "Pass request successfully submitted.",
                    result
                );


                message.className =
                    "pass-form-message " +
                    "pass-form-message-success";


if (
    result &&
    result.attachmentUploaded
) {

    var attachmentCount =
        result.attachmentCount || 0;

    message.innerHTML =
        "<strong>Request submitted successfully.</strong>" +
        "<br>" +
        "The Building Pass Request has been added to the register and " +
        attachmentCount +
        " attachment" +
        (
            attachmentCount === 1
                ? ""
                : "s"
        ) +
        " uploaded successfully.";

}
                else {

                    message.innerHTML =
                        "<strong>Request submitted successfully.</strong>" +
                        "<br>" +
                        "The Building Pass Request has been added to the register.";

                }


                document
                    .getElementById(
                        "passSubmissionForm"
                    )
                    .reset();


                clearPassHolder();


                clearAttachment();


                clearPdfAttachment();


                var requesterField =
                    document.getElementById(
                        "formPassRequester"
                    );


                if (requesterField) {

                    requesterField.value =
                        currentUserEmail ||
                        currentUserDisplayName ||
                        "";

                }


                /*
                ------------------------------------------------
                RESET ACCESS DROPDOWN
                ------------------------------------------------
                */

                updateAccessPermissionsDropdown(
                    true
                );


                submitButton.disabled =
                    false;


                submitButton.textContent =
                    "Submit Request";


                loadData();

            }
        )

        .catch(
            function (error) {

                console.error(
                    "Pass request submission failed.",
                    error
                );


                message.className =
                    "pass-form-message " +
                    "pass-form-message-error";


                message.innerHTML =
                    "<strong>Unable to complete the request.</strong>" +
                    "<br>" +
                    escapeHtml(
                        error.message
                    );


                submitButton.disabled =
                    false;


                submitButton.textContent =
                    "Submit Request";


                if (uploadingContainer) {

                    uploadingContainer.style.display =
                        "none";

                }


                if (uploadingPdfContainer) {

                    uploadingPdfContainer.style.display =
                        "none";

                }

            }
        );

    }


    /*
    ============================================================
    NAVIGATION
    ============================================================
    */

    function bindNavigation() {

        var homeButton =
            document.getElementById(
                "passHomeButton"
            );


        var submitButton =
            document.getElementById(
                "passSubmitButton"
            );


        var dashboardButton =
            document.getElementById(
                "passDashboardButton"
            );


        var adminDashboardButton =
            document.getElementById(
                "passAdminDashboardButton"
            );


        if (homeButton) {

            homeButton.addEventListener(
                "click",
                function () {

                    currentView =
                        "home";

                    renderHome();

                }
            );

        }


        if (submitButton) {

            submitButton.addEventListener(
                "click",
                function () {

                    currentView =
                        "submit";

                    renderSubmissionForm();

                }
            );

        }


        if (dashboardButton) {

            dashboardButton.addEventListener(
                "click",
                function () {

                    currentView =
                        "dashboard";

                    prepareUserDashboard();

                }
            );

        }


        if (adminDashboardButton) {

            adminDashboardButton.addEventListener(
                "click",
                function () {

                    if (!isAdmin) {

                        return;

                    }


                    currentView =
                        "admin";

                    prepareAdminDashboard();

                }
            );

        }

    }


    /*
    ============================================================
    PREPARE USER DASHBOARD
    ============================================================
    */

    function prepareUserDashboard() {

        /*
        --------------------------------------------------------
        SECURITY
        --------------------------------------------------------

        SharePoint item-level permissions determine which
        records were returned by the REST API.

        Therefore we do NOT filter the records here by
        Pass_Requester.

        --------------------------------------------------------
        */

        data =
            allData.slice();


        currentSite =
            "ALL";


        currentStatus =
            "ALL";


        renderDashboard(false);

    }


    /*
    ============================================================
    PREPARE ADMIN DASHBOARD
    ============================================================
    */

    function prepareAdminDashboard() {

        if (!isAdmin) {

            currentView =
                "home";

            renderHome();

            return;

        }


        data =
            allData.slice();


        currentSite =
            "ALL";

        currentStatus =
            "ALL";

        renderDashboard(true);

    }


    /*
    ============================================================
    RENDER DASHBOARD
    ============================================================
    */

    function renderDashboard(adminMode) {

        currentView =
            adminMode
                ? "admin"
                : "dashboard";


        var counts =
            getCounts();


        var filtered =
            getFilteredData();


        var sites =
            getSites();


        dashboard.innerHTML = `

            ${renderNavigation()}


            <div class="dashboard-header">

                <h1>

                    ${
                        adminMode
                            ? "Admin Dashboard"
                            : "Building Pass Request Dashboard"
                    }

                </h1>


                <p>

                    ${
                        adminMode
                            ? "All Building Pass Request"
                            : "Your Building Pass Request"
                    }

                </p>

            </div>


            ${
                adminMode
                    ? `
                        <div class="admin-banner">
                            Administrator view — all pass requests
                            are visible to authorised administrators.
                        </div>
                      `
                    : `
                        <div class="admin-banner">
                            Standard User - You will see all pass requests
                            where you are the pass requester.
                        </div>
                      `
            }


            <div class="summary-grid">


                <div class="summary-card">

                    <div class="summary-label">
                        Total Passes
                    </div>

                    <div class="summary-number number-blue">
                        ${data.length}
                    </div>

                </div>


                <div class="summary-card">

                    <div class="summary-label">
                        Issued
                    </div>

                    <div class="summary-number number-green">
                        ${counts.issued}
                    </div>

                </div>


                <div class="summary-card">

                    <div class="summary-label">
                        Pending
                    </div>

                    <div class="summary-number number-amber">
                        ${counts.pending}
                    </div>

                </div>


                <div class="summary-card">

                    <div class="summary-label">
                        Rejected
                    </div>

                    <div class="summary-number number-red">
                        ${counts.rejected}
                    </div>

                </div>


                <div class="summary-card">

                    <div class="summary-label">
                        Expired
                    </div>

                    <div class="summary-number number-red">
                        ${counts.expired}
                    </div>

                </div>


            </div>


            <div class="controls">


                <input
                    type="text"
                    id="passSearch"
                    class="search-box"
                    placeholder="Search ID, pass holder, site, clearance or requester..."
                    value="${escapeHtml(currentSearch)}"
                    autocomplete="off"
                >


                <select
                    id="passSite"
                    class="filter-select"
                >

                    <option value="ALL">
                        All Sites
                    </option>

                    ${sites.map(
                        function (site) {

                            return `

                                <option
                                    value="${escapeHtml(site)}"
                                    ${
                                        currentSite === site
                                            ? "selected"
                                            : ""
                                    }
                                >
                                    ${escapeHtml(site)}
                                </option>

                            `;

                        }
                    ).join("")}

                </select>


                <select
                    id="passStatus"
                    class="filter-select"
                >

                    <option value="ALL">
                        All Statuses
                    </option>


                    <option
                        value="Issued"
                        ${
                            currentStatus === "Issued"
                                ? "selected"
                                : ""
                        }
                    >
                        Issued
                    </option>


                    <option
                        value="Pending"
                        ${
                            currentStatus === "Pending"
                                ? "selected"
                                : ""
                        }
                    >
                        Pending
                    </option>


                    <option
                        value="Rejected"
                        ${
                            currentStatus === "Rejected"
                                ? "selected"
                                : ""
                        }
                    >
                        Rejected
                    </option>

                </select>


                <div
                    class="record-count"
                    id="passRecordCount"
                >

                    ${filtered.length}
                    of
                    ${data.length}
                    records

                </div>


            </div>


            <div
                class="table-container"
                id="passTableContainer"
            >

                ${renderTable(filtered, adminMode)}

            </div>

        `;


        var searchInput =
            document.getElementById(
                "passSearch"
            );


        if (searchInput) {

            searchInput.addEventListener(
                "input",
                function () {

                    currentSearch =
                        this.value;


                    updateDashboardResults(
                        adminMode
                    );

                }
            );

        }


        var siteSelect =
            document.getElementById(
                "passSite"
            );


        if (siteSelect) {

            siteSelect.addEventListener(
                "change",
                function () {

                    currentSite =
                        this.value;


                    updateDashboardResults(
                        adminMode
                    );

                }
            );

        }


        var statusSelect =
            document.getElementById(
                "passStatus"
            );


        if (statusSelect) {

            statusSelect.addEventListener(
                "change",
                function () {

                    currentStatus =
                        this.value;


                    updateDashboardResults(
                        adminMode
                    );

                }
            );

        }


        bindTableRows();

        bindNavigation();

    }


    /*
    ============================================================
    UPDATE DASHBOARD RESULTS
    ============================================================
    */

    function updateDashboardResults(adminMode) {

        var filtered =
            getFilteredData();


        var tableContainer =
            document.getElementById(
                "passTableContainer"
            );


        var recordCount =
            document.getElementById(
                "passRecordCount"
            );


        if (recordCount) {

            recordCount.textContent =
                filtered.length +
                " of " +
                data.length +
                " records";

        }


        if (tableContainer) {

            tableContainer.innerHTML =
                renderTable(
                    filtered,
                    adminMode
                );

        }


        bindTableRows();

    }


    /*
    ============================================================
    RENDER TABLE
    ============================================================
    */

    function renderTable(
        filtered,
        adminMode
    ) {

        if (filtered.length === 0) {

            return `

                <div class="empty-state">

                    No records match your
                    search or filters.

                </div>

            `;

        }


        return `

            <table class="pass-table">

                <thead>

                    <tr>

                        <th>ID</th>

                        <th>Pass Holder</th>

                        ${
                            adminMode
                                ? "<th>Requester</th>"
                                : ""
                        }

                        <th>Site</th>

                        <th>Clearance</th>

                        <th>Request Date</th>

                        <th>Expiry</th>

<th>Status</th>

${
    adminMode
        ? "<th>Action</th>"
        : ""
}

                    </tr>

                </thead>


                <tbody>

                    ${filtered.map(
                        function (record) {

                            var status =
                                getStatus(record);


                            var expiry =
                                getExpiry(record);


                            return `

                                <tr
                                    data-record-id="${escapeHtml(record.ID)}"
                                >

                                    <td class="id-cell">
                                        #${escapeHtml(record.ID)}
                                    </td>


                                    <td class="holder-cell">
                                        ${escapeHtml(record.Pass_Holder)}
                                    </td>


                                    ${
                                        adminMode
                                        ?
                                        `
                                            <td>
                                                ${escapeHtml(record.Pass_Requester)}
                                            </td>
                                        `
                                        :
                                        ""
                                    }


                                    <td>
                                        ${escapeHtml(record.Site)}
                                    </td>


                                    <td>
                                        ${escapeHtml(record.Clearance_Level)}
                                    </td>


                                    <td>
                                        ${formatDate(record.RequestDate)}
                                    </td>


<td>

    ${formatDate(record.Expiry_Date)}

    <div class="${expiry.className}">
        ${escapeHtml(expiry.label)}
    </div>

</td>


<td>

    <span
        class="status-pill ${status.className}"
    >
        ${escapeHtml(status.label)} -
        ${escapeHtml(status.description)}
    </span>
<!--
    ${
        adminMode &&
        String(record.Status || "").trim() === "Pass with Front Desk Security"
        ?
        `
            <button
                type="button"
                class="issue-pass-button"
                data-issue-pass-id="${escapeHtml(record.ID)}"
            >
                Issue Pass
            </button>
        `
        :
        ""
    }
-->
</td>

${
                                        adminMode &&
                                                String(record.Status || "").trim() === "Pass with Front Desk Security"
        
                                        ?
                                        `
                                            <td>
                                                            <button
                type="button"
                class="issue-pass-button"
                data-issue-pass-id="${escapeHtml(record.ID)}"
            >
                Issue Pass
            </button>
            '
                                            </td>
                                        `
                                        :
                                        ""
                                    }


                                </tr>

                            `;

                        }
                    ).join("")}

                </tbody>

            </table>

        `;

    }


/*
    ============================================================
    BIND TABLE ROWS
    ============================================================
    */

function bindTableRows() {

    var rows =
        dashboard.querySelectorAll(
            "#passTableContainer tbody tr"
        );


    rows.forEach(
        function (row) {

            // ----------------------------------------------------
            // Row click - open details
            // ----------------------------------------------------

            row.addEventListener(
                "click",
                function (event) {

                    // Don't open the details panel when the
                    // Issue Pass button has been clicked.
                    if (
                        event.target.closest(
                            ".issue-pass-button"
                        )
                    ) {
                        return;
                    }


                    var id =
                        this.getAttribute(
                            "data-record-id"
                        );


                    var record =
                        data.find(
                            function (item) {

                                return String(
                                    item.ID
                                ) ===
                                String(id);

                            }
                        );


                    if (record) {

                        showDetails(
                            record
                        );

                    }

                }
            );


            // ----------------------------------------------------
            // Issue Pass button
            // ----------------------------------------------------

            var issueButton =
                row.querySelector(
                    ".issue-pass-button"
                );


            if (issueButton) {

                issueButton.addEventListener(
                    "click",
                    function (event) {

                        event.stopPropagation();


                        var id =
                            this.getAttribute(
                                "data-issue-pass-id"
                            );


                        issuePass(id);

                    }
                );

            }

        }
    );

}

/*
============================================================
ISSUE PASS
============================================================
*/

async function issuePass(itemId) {

    if (
        !confirm(
            "Are you sure you want to issue this pass?"
        )
    ) {
        return;
    }


    try {

        /*
        ========================================================
        GET LOGGED-IN USER
        ========================================================
        */

        var currentUserResponse =
            await fetch(
                SITE_URL +
                "/_api/web/currentuser",
                {
                    method: "GET",

                    headers: {
                        "Accept":
                            "application/json;odata=verbose"
                    }
                }
            );


        if (!currentUserResponse.ok) {

            throw new Error(
                "Unable to identify the logged-in user."
            );

        }


        var currentUserData =
            await currentUserResponse.json();


        var loggedInUser =
            currentUserData.d.Title ||
            currentUserData.d.LoginName ||
            "Unknown user";


        /*
        ========================================================
        GET REQUEST DIGEST
        ========================================================
        */

        var contextResponse =
            await fetch(
                SITE_URL +
                "/_api/contextinfo",
                {
                    method: "POST",

                    headers: {
                        "Accept":
                            "application/json;odata=verbose"
                    }
                }
            );


        if (!contextResponse.ok) {

            var contextError =
                await contextResponse.text();

            console.error(
                "Context info error:",
                contextError
            );

            throw new Error(
                "Unable to obtain the SharePoint request digest."
            );

        }


        var contextData =
            await contextResponse.json();


        var requestDigest =
            contextData.d
                .GetContextWebInformation
                .FormDigestValue;


        /*
        ========================================================
        TODAY'S DATE
        ========================================================
        */

        var today =
            new Date();


        var day =
            String(
                today.getDate()
            ).padStart(2, "0");


        var month =
            String(
                today.getMonth() + 1
            ).padStart(2, "0");


        var year =
            today.getFullYear();


        var formattedDate =
            `${day}/${month}/${year}`;


        /*
        ========================================================
        NEW STATUS
        ========================================================
        */

        var newStatus =
            `Pass issued (${formattedDate}) by ${loggedInUser}`;


        /*
        ========================================================
        UPDATE SHAREPOINT ITEM
        ========================================================
        */

        var updateResponse =
            await fetch(

                LIST_API_URL +
                "(" +
                itemId +
                ")",

                {
                    method: "POST",

                    headers: {

                        "Accept":
                            "application/json;odata=verbose",

                        "Content-Type":
                            "application/json;odata=verbose",

                        "X-HTTP-Method":
                            "MERGE",

                        "IF-MATCH":
                            "*",

                        "X-RequestDigest":
                            requestDigest

                    },

                    body: JSON.stringify({

                        "__metadata": {

                            "type":
                                LIST_ITEM_TYPE

                        },

                        "Status":
                            newStatus

                    })

                }

            );


        /*
        ========================================================
        CHECK RESPONSE
        ========================================================
        */

        var responseText =
            await updateResponse.text();


        if (!updateResponse.ok) {

            console.error(
                "SharePoint update failed:",
                updateResponse.status,
                responseText
            );

            throw new Error(
                `SharePoint returned ${updateResponse.status}: ${responseText}`
            );

        }


        /*
        ========================================================
        SUCCESS
        ========================================================
        */

        alert(
            "Pass issued successfully."
        );


        location.reload();


    } catch (error) {

        console.error(
            "Issue Pass error:",
            error
        );


        alert(
            "Unable to issue the pass.\n\n" +
            error.message
        );

    }

}
    /*
    ============================================================
    DETAILS MODAL
    ============================================================
    */

    function showDetails(record) {

        var status =
            getStatus(record);


        var expiry =
            getExpiry(record);


        var overlay =
            document.createElement(
                "div"
            );


        overlay.className =
            "details-overlay";


        overlay.innerHTML = `

            <div class="details-modal">


                <div class="modal-header">

                    <h2>
                        Pass Request #${escapeHtml(record.ID)}
                    </h2>


                    <button
                        class="close-button"
                        type="button"
                    >
                        ×
                    </button>

                </div>


                <div class="modal-body">


                    <div class="detail-grid">


                        <div class="detail-section">

                            <div class="detail-section-title">
                                Pass Details
                            </div>


                            <div class="detail-section-body">


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Status
                                    </div>

                                    <div class="detail-value">

                                        <span
                                            class="status-pill ${status.className}"
                                        >
                                            ${escapeHtml(status.label)}
                                        </span>

                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Pass Holder
                                    </div>

                                    <div class="detail-value">
                                        ${escapeHtml(record.Pass_Holder)}
                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        National Insurance Number
                                    </div>

                                    <div class="detail-value">
                                        ${escapeHtml(record.NI)}
                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Site
                                    </div>

                                    <div class="detail-value">
                                        ${escapeHtml(record.Site)}
                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Clearance
                                    </div>

                                    <div class="detail-value">
                                        ${escapeHtml(record.Clearance_Level)}
                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Request Date
                                    </div>

                                    <div class="detail-value">
                                        ${formatDate(record.RequestDate)}
                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Required Date
                                    </div>

                                    <div class="detail-value">
                                        ${formatDate(record.Required_Date)}
                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Expiry Date
                                    </div>

                                    <div class="detail-value">

                                        ${formatDate(record.Expiry_Date)}

                                        <div class="${expiry.className}">
                                            ${escapeHtml(expiry.label)}
                                        </div>

                                    </div>

                                </div>


                            </div>

                        </div>


                        <div class="detail-section">

                            <div class="detail-section-title">
                                Pass Requester
                            </div>


                            <div class="detail-section-body">


                                ${
                                    record.Author &&
                                    record.Author.Picture

                                    ?

                                    `

                                        <div class="person">

                                            <img
                                                src="${escapeHtml(record.Author.Picture)}"
                                                alt=""
                                            >

                                            <div>

                                                <div class="person-name">
                                                    ${escapeHtml(record.Author.DisplayName)}
                                                </div>

                                                <div class="person-email">
                                                    ${escapeHtml(record.Author.Email)}
                                                </div>

                                            </div>

                                        </div>

                                    `

                                    :

                                    ""

                                }


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Requester
                                    </div>

                                    <div class="detail-value">
                                        ${escapeHtml(record.Pass_Requester)}
                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Department
                                    </div>

                                    <div class="detail-value">

                                        ${
                                            record.Author
                                                ? escapeHtml(
                                                    record.Author.Department
                                                )
                                                : ""
                                        }

                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Job Title
                                    </div>

                                    <div class="detail-value">

                                        ${
                                            record.Author
                                                ? escapeHtml(
                                                    record.Author.JobTitle
                                                )
                                                : ""
                                        }

                                    </div>

                                </div>


                            </div>

                        </div>


                        <div class="detail-section detail-section-full">

                            <div class="detail-section-title">
                                Access Permissions
                            </div>


                            <div class="detail-section-body">

                                <div class="access-tags">

                                    ${accessTags(record.Access)}

                                </div>

                            </div>

                        </div>


                        <div class="detail-section detail-section-full">

                            <div class="detail-section-title">
                                Justification
                            </div>


                            <div class="detail-section-body">

                                <div class="justification">

                                    ${escapeHtml(
                                        record.Justification
                                    )}

                                </div>

                            </div>

                        </div>


                        <div class="detail-section detail-section-full">

                            <div class="detail-section-title">
                                Record Information
                            </div>


                            <div class="detail-section-body">


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Record ID
                                    </div>

                                    <div class="detail-value">
                                        ${escapeHtml(record.ID)}
                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Created
                                    </div>

                                    <div class="detail-value">
                                        ${formatDateTime(record.Created)}
                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Modified
                                    </div>

                                    <div class="detail-value">
                                        ${formatDateTime(record.Modified)}
                                    </div>

                                </div>


                                <div class="detail-row">

                                    <div class="detail-label">
                                        Status Detail
                                    </div>

                                    <div class="detail-value">
                                        ${escapeHtml(record.Status)}
                                    </div>

                                </div>


                            </div>

                        </div>


                    </div>

                </div>

            </div>

        `;


        document.body.appendChild(
            overlay
        );


        overlay
            .querySelector(
                ".close-button"
            )
            .addEventListener(
                "click",
                function () {

                    overlay.remove();

                }
            );


        overlay.addEventListener(
            "click",
            function (event) {

                if (
                    event.target ===
                    overlay
                ) {

                    overlay.remove();

                }

            }
        );


        function closeOnEscape(event) {

            if (
                event.key ===
                "Escape"
            ) {

                overlay.remove();


                document.removeEventListener(
                    "keydown",
                    closeOnEscape
                );

            }

        }


        document.addEventListener(
            "keydown",
            closeOnEscape
        );

    }


    /*
    ============================================================
    LOAD DATA FROM SHAREPOINT
    ============================================================
    */

    function loadData() {

        console.log(
            "Pass Dashboard: Loading SharePoint data..."
        );


        var url =
            LIST_API_URL +
            "?$orderby=ID%20desc";


        console.log(
            "SharePoint API URL:",
            url
        );


        fetch(
            url,
            {
                method: "GET",

                credentials:
                    "same-origin",

                headers: {

                    "Accept":
                        "application/json;odata=verbose"

                },

                cache:
                    "no-store"

            }
        )

        .then(
            function (response) {

                if (!response.ok) {

                    return response.text()
                        .then(
                            function (text) {

                                console.error(
                                    "SharePoint API error response:",
                                    text
                                );


                                throw new Error(
                                    "SharePoint returned HTTP " +
                                    response.status +
                                    " - " +
                                    text
                                );

                            }
                        );

                }


                return response.json();

            }
        )

        .then(
            function (json) {

                console.log(
                    "SharePoint response:",
                    json
                );


                if (
                    !json ||
                    !json.d ||
                    !Array.isArray(
                        json.d.results
                    )
                ) {

                    throw new Error(
                        "SharePoint returned an unexpected response format."
                    );

                }


                allData =
                    json.d.results;


                console.log(
                    "Total SharePoint records loaded:",
                    allData.length
                );


                if (
                    currentView ===
                    "admin"
                ) {

                    if (isAdmin) {

                        prepareAdminDashboard();

                    }
                    else {

                        currentView =
                            "home";

                        renderHome();

                    }

                }

                else if (
                    currentView ===
                    "dashboard"
                ) {

                    prepareUserDashboard();

                }

                else if (
                    currentView ===
                    "submit"
                ) {

                    return;

                }

                else {

                    renderHome();

                }

            }
        )

        .catch(
            function (error) {

                console.error(
                    "Pass Dashboard: Unable to load SharePoint data.",
                    error
                );


                dashboard.innerHTML = `

                    <div class="error-state">

                        <h2>
                            Unable to load pass data
                        </h2>

                        <p>
                            The dashboard could not retrieve
                            your Building Pass records from SharePoint.
                        </p>

                        <p>
                            <strong>Error:</strong>
                            ${escapeHtml(
                                error.message
                            )}
                        </p>

                        <p>
                            Please check that you have access
                            to the Building Pass register.
                        </p>

                    </div>

                `;

            }
        );

    }


    /*
    ============================================================
    START APPLICATION
    ============================================================
    */

    Promise.all(
        [
            getCurrentUser(),
            getAvailableSites()
        ]
    )
    .then(
        function () {

            console.log(
                "SharePoint user and Site list loaded."
            );


            renderHome();


            loadData();

        }
    )
    .catch(
        function (error) {

            console.error(
                "Pass application initialisation failed.",
                error
            );


            renderHome();


            loadData();

        }
    );


    /*
    ============================================================
    AUTOMATIC REFRESH
    ============================================================

    Refresh every 5 minutes.

    ============================================================
    */

    setInterval(
        function () {

            loadData();

        },
        5 * 60 * 1000
    );


    console.log(
        "Pass Dashboard initialised."
    );


})();
