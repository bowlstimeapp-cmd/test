javascript:
(function () {

    "use strict";

    console.log("NSV Dashboard JS started");


    /* ============================================================
       DASHBOARD CONTAINER
       ============================================================ */

    var dashboardContainer =
        document.getElementById("passDashboard");


    if (!dashboardContainer) {

        console.error(
            "NSV Dashboard: passDashboard was not found."
        );

        return;

    }


    /* ============================================================
       CONFIGURATION
       ============================================================ */


var dashboardMode = false;
var dashboardPeriod = 1;


    var SITE_URL =
        "https://mcga.sharepoint.com/sites/InformationAssurance";


    var LIST_TITLE =
        "NSV Workbook";


    var APPLICATIONS_LIST_TITLE =
        "NSV Applications";


    var MAX_ATTACHMENT_SIZE =
        25 * 1024 * 1024;


    var TABLE_COLUMNS = [

        {
            field: "field_6",
            header: "Full Name",
            type: "text"
        },

        {
            field: "field_2",
            header: "Organisation",
            type: "text"
        },

        {
            field: "field_5",
            header: "Directorate",
            type: "text"
        },

        {
            field: "Role",
            header: "Role",
            type: "text"
        },

        {
            field: "field_10",
            header: "Email Address",
            type: "text"
        },

        {
            field: "field_14",
            header: "DOB",
            type: "date"
        },

        {
            field: "field_13",
            header: "Clearance Level",
            type: "status"
        },

        {
            field: "field_15",
            header: "Valid From",
            type: "date"
        },

        {
            field: "field_16",
            header: "Valid To",
            type: "date"
        },

        {
            field: "field_9",
            header: "National Insurance",
            type: "text"
        },

        {
            field: "field_4",
            header: "Leaving Date",
            type: "date"
        }

    ];


    /* ============================================================
       STATE
       ============================================================ */

    var allItems = [];

    var filteredItems = [];

    var columns =
        TABLE_COLUMNS.slice();

    var activeFilter = null;

    var updateMode = false;

    var selectedItem = null;

    var columnWidths = {};

    var sortState = {
        field: null,
        direction: "asc"
    };


    /* ============================================================
       APPLICATIONS STATE
       ============================================================ */

    var allApplicationItems = [];

    var filteredApplicationItems = [];

    var applicationsMode = false;

    var selectedApplicationItem = null;

    var applicationSortState = {
        field: null,
        direction: "asc"
    };


    /* ============================================================
       TABLE CONTAINER RESIZING
       ============================================================ */

    var panelResizeState = null;


    function attachPanelResizeHandler() {

        var handle =
            document.querySelector(
                ".nsv-panel-resize-handle"
            );


        if (!handle) {

            return;

        }


        handle.addEventListener(
            "mousedown",
            startPanelResize
        );

    }


    function startPanelResize(event) {

        event.preventDefault();
        event.stopPropagation();


        var panel =
            document.querySelector(
                ".nsv-panel-resizable"
            );


        if (!panel) {

            return;

        }


        panelResizeState = {

            startX:
                event.clientX,

            startWidth:
                panel.getBoundingClientRect().width

        };


        document.body.style.cursor =
            "col-resize";

        document.body.style.userSelect =
            "none";


        document.addEventListener(
            "mousemove",
            handlePanelResize
        );

        document.addEventListener(
            "mouseup",
            stopPanelResize
        );

    }


    function handlePanelResize(event) {

        if (!panelResizeState) {

            return;

        }


        var panel =
            document.querySelector(
                ".nsv-panel-resizable"
            );


        if (!panel) {

            return;

        }


        var delta =
            event.clientX -
            panelResizeState.startX;


        var newWidth =
            panelResizeState.startWidth +
            delta;


        newWidth =
            Math.max(
                700,
                Math.min(
                    1800,
                    newWidth
                )
            );


        panel.style.width =
            Math.round(
                newWidth
            ) +
            "px";

    }


    function stopPanelResize() {

        panelResizeState =
            null;


        document.body.style.cursor =
            "";

        document.body.style.userSelect =
            "";


        document.removeEventListener(
            "mousemove",
            handlePanelResize
        );

        document.removeEventListener(
            "mouseup",
            stopPanelResize
        );

    }


    /* ============================================================
       PAGE HTML
       ============================================================ */

    dashboardContainer.innerHTML = `

        <div class="nsv-dashboard">

            <div class="nsv-header">

                <div>

                    <h1 class="nsv-title">
                        NSV Workbook Dashboard
                    </h1>

                    <p class="nsv-subtitle">
                        Information Assurance
                    </p>

                </div>

                <button
                    id="nsv-refresh"
                    class="nsv-button"
                    type="button"
                >
                    Refresh
                </button>

            </div>


            <div
                id="nsv-status"
                class="nsv-status"
            >
                Connecting to SharePoint...
            </div>


            <div class="nsv-kpi-grid">

                <div class="nsv-kpi-card">

                    <div class="nsv-kpi-label">
                        Total Records
                    </div>

                    <div
                        id="nsv-total"
                        class="nsv-kpi-value"
                    >
                        —
                    </div>

                    <div class="nsv-kpi-detail">
                        All clearance holders
                    </div>

                    <div
                        id="nsv-total-expired"
                        class="nsv-kpi-footer"
                    >
                        Total Expired: —
                    </div>

                </div>


                <div
                    id="nsv-mca-kpi"
                    class="nsv-kpi-card nsv-kpi-clickable"
                    role="button"
                    tabindex="0"
                    title="Filter table to MCA personnel"
                >

                    <div class="nsv-kpi-label">
                        MCA
                    </div>

                    <div
                        id="nsv-mca"
                        class="nsv-kpi-value"
                    >
                        —
                    </div>

                    <div class="nsv-kpi-detail">
                        Click to filter by Organisation
                    </div>

                    <div
                        id="nsv-mca-expired"
                        class="nsv-kpi-footer"
                    >
                        MCA Expired: —
                    </div>

                </div>


                <div
                    id="nsv-contractor-kpi"
                    class="nsv-kpi-card nsv-kpi-clickable"
                    role="button"
                    tabindex="0"
                    title="Filter table to contractors"
                >

                    <div class="nsv-kpi-label">
                        Contractors
                    </div>

                    <div
                        id="nsv-contractors"
                        class="nsv-kpi-value"
                    >
                        —
                    </div>

                    <div class="nsv-kpi-detail">
                        Click to filter by Organisation
                    </div>

                    <div
                        id="nsv-contractor-expired"
                        class="nsv-kpi-footer"
                    >
                        Contractor Expired: —
                    </div>

                </div>


                <div class="nsv-kpi-card">

                    <div class="nsv-kpi-label">
                        Last Modified
                    </div>

                    <div
                        id="nsv-modified"
                        class="nsv-kpi-value nsv-kpi-date"
                    >
                        —
                    </div>

                    <div class="nsv-kpi-detail">
                        NSV Workbook
                    </div>

                </div>

            </div>


            <div class="nsv-actions">

                <button
                    id="nsv-add-clearance"
                    class="nsv-action-button nsv-primary-button"
                    type="button"
                >

                    <span
                        class="nsv-action-icon"
                        aria-hidden="true"
                    >
                        +
                    </span>

                    Add new clearance holder

                </button>


                <button
                    id="nsv-update-clearance"
                    class="nsv-action-button"
                    type="button"
                    aria-pressed="false"
                >

                    <span
                        class="nsv-action-icon"
                        aria-hidden="true"
                    >
                        ✎
                    </span>

                    Update a clearance

                </button>


                <button
                    id="nsv-expired-clearance"
                    class="nsv-action-button"
                    type="button"
                    aria-pressed="false"
                >

                    <span
                        class="nsv-action-icon"
                        aria-hidden="true"
                    >
                        !
                    </span>

                    Filter by expired

                </button>


                <button
                    id="nsv-renewal-filter"
                    class="nsv-action-button"
                    type="button"
                    aria-pressed="false"
                >

                    <span
                        class="nsv-action-icon"
                        aria-hidden="true"
                    >
                        ↻
                    </span>

                    Filter by Renewal

                </button>

                <button
    id="nsv-dashboards"
    class="nsv-action-button"
    type="button"
    aria-pressed="false"
>
    <span class="nsv-action-icon" aria-hidden="true">▤</span>
    Dashboards
</button>

                <button
    id="nsv-applications"
    class="nsv-action-button"
    type="button"
    aria-pressed="false"
>
    <span class="nsv-action-icon" aria-hidden="true">➜</span>
    Applications
</button>

            </div>


            <div class="nsv-panel nsv-panel-resizable">

                <div
                    class="nsv-panel-resize-handle"
                    title="Drag to resize table width"
                    aria-label="Resize table width"
                ></div>


                <div class="nsv-panel-header">

                    <div>

                        <h2 class="nsv-panel-title">
                            NSV Workbook
                        </h2>

                        <p
                            id="nsv-panel-description"
                            class="nsv-panel-description"
                        >
                            SharePoint list records
                        </p>

                    </div>


                    <input
                        id="nsv-search"
                        class="nsv-search"
                        type="search"
                        placeholder="Search records..."
                        autocomplete="off"
                    >

                </div>


                <div
                    id="nsv-update-hint"
                    class="nsv-update-hint"
                    style="display:none;"
                >
                    Update mode is active. Click a row to edit that
                    clearance holder.
                </div>


                <div
                    id="nsv-table-container"
                    class="nsv-table-container"
                >

                    <div class="nsv-loading">
                        Loading records...
                    </div>

                </div>

            </div>

        </div>

    `;


    /* ============================================================
       BASIC HELPERS
       ============================================================ */

    function byId(id) {

        return document.getElementById(id);

    }


    function setStatus(message) {

        var element =
            byId("nsv-status");


        if (element) {

            element.textContent =
                message;

        }

    }


    function getFieldValue(
        item,
        field
    ) {

        if (!item) {

            return "";

        }


        var value =
            item[field];


        if (
            value === null ||
            value === undefined
        ) {

            return "";

        }


        return String(value);

    }


    /* ============================================================
       DATE / EXPIRY / RENEWAL HELPERS
       ============================================================ */

    function getDateOnly(value) {

        if (!value) {

            return null;

        }


        var date =
            new Date(value);


        if (
            isNaN(
                date.getTime()
            )
        ) {

            return null;

        }


        date.setHours(
            0,
            0,
            0,
            0
        );


        return date;

    }


    function isExpired(item) {

        var validTo =
            getFieldValue(
                item,
                "field_16"
            );


        if (!validTo) {

            return false;

        }


        var expiryDate =
            getDateOnly(
                validTo
            );


        if (!expiryDate) {

            return false;

        }


        var today =
            new Date();


        today.setHours(
            0,
            0,
            0,
            0
        );


        return expiryDate < today;

    }


    function isDueForRenewal(item) {

        var validTo =
            getFieldValue(
                item,
                "field_16"
            );


        if (!validTo) {

            return false;

        }


        var expiryDate =
            getDateOnly(
                validTo
            );


        if (!expiryDate) {

            return false;

        }


        var today =
            new Date();


        today.setHours(
            0,
            0,
            0,
            0
        );


        var renewalDate =
            new Date(today);


        renewalDate.setDate(
            renewalDate.getDate() + 90
        );


        return (
            expiryDate >= today &&
            expiryDate <= renewalDate
        );

    };

    function getDailyCreatedCounts() {

    var now = new Date();
    var year = now.getFullYear();
    var month = now.getMonth();

    var daysInMonth =
        new Date(year, month + 1, 0).getDate();

    var counts =
        new Array(daysInMonth).fill(0);


    allItems.forEach(function (item) {

        var created =
            item.Created
                ? new Date(item.Created)
                : null;


        if (!created || isNaN(created.getTime())) {
            return;
        }


        if (
            created.getFullYear() === year &&
            created.getMonth() === month
        ) {
            counts[created.getDate() - 1] += 1;
        }

    });


    var labels =
        counts.map(function (_, index) {
            return String(index + 1);
        });


    return { labels: labels, values: counts };

}


function getMonthlyCreatedCounts(monthCount) {

    var now = new Date();
    var buckets = [];


    for (var i = monthCount - 1; i >= 0; i--) {

        var bucketDate =
            new Date(now.getFullYear(), now.getMonth() - i, 1);


        buckets.push({
            year: bucketDate.getFullYear(),
            month: bucketDate.getMonth(),
            label: bucketDate.toLocaleDateString(
                "en-GB",
                { month: "short", year: "numeric" }
            ),
            count: 0
        });

    }


    allItems.forEach(function (item) {

        var created =
            item.Created
                ? new Date(item.Created)
                : null;


        if (!created || isNaN(created.getTime())) {
            return;
        }


        buckets.forEach(function (bucket) {

            if (
                created.getFullYear() === bucket.year &&
                created.getMonth() === bucket.month
            ) {
                bucket.count += 1;
            }

        });

    });


    return {
        labels: buckets.map(function (b) { return b.label; }),
        values: buckets.map(function (b) { return b.count; })
    };

}

    


    /* ============================================================
       SHAREPOINT API
       ============================================================ */

    async function apiGet(url) {

        console.log(
            "SharePoint REST GET:",
            url
        );


        var response =
            await fetch(
                url,
                {
                    method: "GET",
                    credentials: "same-origin",
                    headers: {
                        "Accept":
                            "application/json;odata=nometadata"
                    }
                }
            );


        if (!response.ok) {

            var text = "";

            try {

                text =
                    await response.text();

            }
            catch (e) {

                text = "";

            }


            throw new Error(
                "SharePoint API error: " +
                response.status +
                " " +
                response.statusText +
                "\n" +
                text
            );

        }


        return await response.json();

    }


    async function apiPost(
        url,
        body,
        extraHeaders
    ) {

        var headers = {

            "Accept":
                "application/json;odata=nometadata"

        };


        if (extraHeaders) {

            Object.keys(
                extraHeaders
            ).forEach(
                function (key) {

                    headers[key] =
                        extraHeaders[key];

                }
            );

        }


        var response =
            await fetch(
                url,
                {
                    method: "POST",
                    credentials: "same-origin",
                    headers: headers,
                    body: body
                }
            );


        if (!response.ok) {

            var text = "";

            try {

                text =
                    await response.text();

            }
            catch (e) {

                text = "";

            }


            throw new Error(
                "SharePoint API error: " +
                response.status +
                " " +
                response.statusText +
                "\n" +
                text
            );

        }


        var responseText =
            await response.text();


        if (!responseText) {

            return {};

        }


        try {

            return JSON.parse(
                responseText
            );

        }
        catch (e) {

            return {};

        }

    }


    async function apiMerge(
        url,
        body,
        extraHeaders
    ) {

        var headers = {

            "Accept":
                "application/json;odata=nometadata",

            "Content-Type":
                "application/json;odata=nometadata",

            "IF-MATCH":
                "*",

            "X-HTTP-Method":
                "MERGE"

        };


        if (extraHeaders) {

            Object.keys(
                extraHeaders
            ).forEach(
                function (key) {

                    headers[key] =
                        extraHeaders[key];

                }
            );

        }


        var response =
            await fetch(
                url,
                {
                    method: "POST",
                    credentials: "same-origin",
                    headers: headers,
                    body: JSON.stringify(body)
                }
            );


        if (!response.ok) {

            var text = "";

            try {

                text =
                    await response.text();

            }
            catch (e) {

                text = "";

            }


            throw new Error(
                "SharePoint API error: " +
                response.status +
                " " +
                response.statusText +
                "\n" +
                text
            );

        }


        return true;

    }


    async function apiDelete(
        url,
        extraHeaders
    ) {

        var headers = {

            "Accept":
                "application/json;odata=nometadata",

            "IF-MATCH":
                "*"

        };


        if (extraHeaders) {

            Object.keys(
                extraHeaders
            ).forEach(
                function (key) {

                    headers[key] =
                        extraHeaders[key];

                }
            );

        }


        var response =
            await fetch(
                url,
                {
                    method: "POST",
                    credentials: "same-origin",
                    headers: {
                        "Accept":
                            headers["Accept"],

                        "IF-MATCH":
                            headers["IF-MATCH"],

                        "X-RequestDigest":
                            headers["X-RequestDigest"],

                        "X-HTTP-Method":
                            "DELETE"
                    }
                }
            );


        if (!response.ok) {

            var text = "";

            try {

                text =
                    await response.text();

            }
            catch (e) {

                text = "";

            }


            throw new Error(
                "SharePoint API error: " +
                response.status +
                " " +
                response.statusText +
                "\n" +
                text
            );

        }


        return true;

    }


    /* ============================================================
       REQUEST DIGEST
       ============================================================ */

    async function getRequestDigest() {

        var url =
            SITE_URL +
            "/_api/contextinfo";


        var response =
            await fetch(
                url,
                {
                    method: "POST",
                    credentials: "same-origin",
                    headers: {
                        "Accept":
                            "application/json;odata=verbose",

                        "Content-Type":
                            "application/json;odata=verbose"
                    },
                    body: ""
                }
            );


        if (!response.ok) {

            var text = "";

            try {

                text =
                    await response.text();

            }
            catch (e) {

                text = "";

            }


            throw new Error(
                "Unable to obtain SharePoint request digest.\n" +
                response.status +
                " " +
                response.statusText +
                "\n" +
                text
            );

        }


        var data =
            await response.json();


        if (
            data &&
            data.d &&
            data.d.GetContextWebInformation
        ) {

            return data
                .d
                .GetContextWebInformation
                .FormDigestValue;

        }


        if (
            data &&
            data.GetContextWebInformation
        ) {

            return data
                .GetContextWebInformation
                .FormDigestValue;

        }


        throw new Error(
            "SharePoint request digest was not returned."
        );

    }


    /* ============================================================
       PDF UPLOAD
       ============================================================ */

    async function uploadPdfAttachment(
        itemId,
        file
    ) {

        if (!file) {

            return null;

        }


        var fileName =
            String(
                file.name || ""
            );


        var extension =
            fileName
                .split(".")
                .pop()
                .toLowerCase();


        if (extension !== "pdf") {

            throw new Error(
                "The clearance document must be a PDF file."
            );

        }


        if (
            file.type &&
            file.type !== "application/pdf"
        ) {

            throw new Error(
                "The selected clearance document is not recognised as a PDF."
            );

        }


        if (
            file.size >
            MAX_ATTACHMENT_SIZE
        ) {

            throw new Error(
                "The PDF is too large. The maximum permitted size is 25 MB."
            );

        }


        setStatus(
            "Uploading clearance PDF..."
        );


        var digest =
            await getRequestDigest();


        var safeFileName =
            fileName.replace(
                /[\\/:*?"<>|#%]/g,
                "_"
            );


        var attachmentUrl =

            SITE_URL +

            "/_api/web/lists/GetByTitle('" +

            encodeURIComponent(
                LIST_TITLE
            ) +

            "')/items(" +

            itemId +

            ")/AttachmentFiles/add(FileName='" +

            encodeURIComponent(
                safeFileName
            ) +

            "')";


        var response =
            await fetch(
                attachmentUrl,
                {
                    method: "POST",
                    credentials: "same-origin",

                    headers: {

                        "Accept":
                            "application/json;odata=nometadata",

                        "X-RequestDigest":
                            digest,

                        "Content-Type":
                            "application/pdf"

                    },

                    body: file
                }
            );


        if (!response.ok) {

            var text = "";

            try {

                text =
                    await response.text();

            }
            catch (e) {

                text = "";

            }


            throw new Error(
                "PDF upload failed: " +
                response.status +
                " " +
                response.statusText +
                "\n" +
                text
            );

        }


        return true;

    }


    /* ============================================================
       LOAD LIST
       ============================================================ */

    async function loadList() {

        setStatus(
            "Reading NSV Workbook..."
        );


        var listUrl =

            SITE_URL +

            "/_api/web/lists/GetByTitle('" +

            encodeURIComponent(
                LIST_TITLE
            ) +

            "')?$select=Title,ItemCount,LastItemModifiedDate";


        var list =
            await apiGet(
                listUrl
            );


        byId(
            "nsv-modified"
        ).textContent =
            formatDate(
                list.LastItemModifiedDate
            );

    }


    /* ============================================================
       LOAD ITEMS
       ============================================================ */

    async function loadItems() {

        setStatus(
            "Loading NSV Workbook records..."
        );


        var url =

            SITE_URL +

            "/_api/web/lists/GetByTitle('" +

            encodeURIComponent(
                LIST_TITLE
            ) +

            "')/items?$top=5000";


        var result =
            await apiGet(
                url
            );


        allItems =
            result.value || [];


applyCurrentFilter();
updateKpis();

if (dashboardMode) {
    renderDashboard();
} else if (applicationsMode) {
    /*
     * The Applications tab is currently on screen. The
     * Workbook data (and KPIs) just refreshed in the
     * background, but the visible table should stay on
     * Applications rather than being replaced here.
     */
} else {
    renderTable();
}

updatePanelDescription();


        if (!applicationsMode) {

            setStatus(
                allItems.length +
                " records loaded successfully."
            );

        }

    }


    /* ============================================================
       KPI CALCULATIONS
       ============================================================ */

    function updateKpis() {

        var total =
            allItems.length;


        var totalExpired =
            allItems.filter(
                function (item) {

                    return isExpired(
                        item
                    );

                }
            ).length;


        var mcaItems =
            allItems.filter(
                function (item) {

                    return getFieldValue(
                        item,
                        "field_2"
                    )
                        .trim()
                        .toLowerCase() ===
                        "mca";

                }
            );


        var mcaCount =
            mcaItems.length;


        var mcaExpired =
            mcaItems.filter(
                function (item) {

                    return isExpired(
                        item
                    );

                }
            ).length;


        var contractorItems =
            allItems.filter(
                function (item) {

                    return getFieldValue(
                        item,
                        "field_2"
                    )
                        .toLowerCase()
                        .includes(
                            "contractor"
                        );

                }
            );


        var contractorCount =
            contractorItems.length;


        var contractorExpired =
            contractorItems.filter(
                function (item) {

                    return isExpired(
                        item
                    );

                }
            ).length;


        byId(
            "nsv-total"
        ).textContent =
            total.toLocaleString(
                "en-GB"
            );


        byId(
            "nsv-mca"
        ).textContent =
            mcaCount.toLocaleString(
                "en-GB"
            );


        byId(
            "nsv-contractors"
        ).textContent =
            contractorCount.toLocaleString(
                "en-GB"
            );


        byId(
            "nsv-total-expired"
        ).textContent =
            "Total Expired: " +
            totalExpired.toLocaleString(
                "en-GB"
            );


        byId(
            "nsv-mca-expired"
        ).textContent =
            "MCA Expired: " +
            mcaExpired.toLocaleString(
                "en-GB"
            );


        byId(
            "nsv-contractor-expired"
        ).textContent =
            "Contractor Expired: " +
            contractorExpired.toLocaleString(
                "en-GB"
            );

    }


    /* ============================================================
       FILTER BUTTON STATE
       ============================================================ */

    function clearActiveKpiClasses() {

        var mca =
            byId(
                "nsv-mca-kpi"
            );


        var contractor =
            byId(
                "nsv-contractor-kpi"
            );


        if (mca) {

            mca.classList.remove(
                "nsv-kpi-active"
            );

        }


        if (contractor) {

            contractor.classList.remove(
                "nsv-kpi-active"
            );

        }

    }


    function updateExpiredButton() {

        var button =
            byId(
                "nsv-expired-clearance"
            );


        if (!button) {

            return;

        }


        var active =
            activeFilter === "expired";


        button.classList.toggle(
            "nsv-update-active",
            active
        );


        button.setAttribute(
            "aria-pressed",
            active ? "true" : "false"
        );

    }


    function updateRenewalButton() {

        var button =
            byId(
                "nsv-renewal-filter"
            );


        if (!button) {

            return;

        }


        var active =
            activeFilter === "renewal";


        button.classList.toggle(
            "nsv-filter-active",
            active
        );


        button.classList.toggle(
            "nsv-update-active",
            active
        );


        button.setAttribute(
            "aria-pressed",
            active ? "true" : "false"
        );

    }


    /* ============================================================
       SEARCH MATCHING
       ============================================================ */

    function itemMatchesSearch(
        item,
        searchText
    ) {

        return TABLE_COLUMNS.some(
            function (column) {

                var value =
                    item[
                        column.field
                    ];


                return String(
                    value === null ||
                    value === undefined
                        ? ""
                        : value
                )
                    .toLowerCase()
                    .includes(
                        searchText
                    );

            }
        );

    }


    /* ============================================================
       FILTER LOGIC
       ============================================================ */

    function applyCurrentFilter() {

        var searchInput =
            byId(
                "nsv-search"
            );


        var searchText =
            searchInput
                ? searchInput.value
                    .toLowerCase()
                    .trim()
                : "";


        clearActiveKpiClasses();


        if (
            activeFilter === "mca"
        ) {

            byId(
                "nsv-mca-kpi"
            ).classList.add(
                "nsv-kpi-active"
            );


            filteredItems =
                allItems.filter(
                    function (item) {

                        var organisation =
                            getFieldValue(
                                item,
                                "field_2"
                            )
                                .trim()
                                .toLowerCase();


                        return (
                            organisation ===
                            "mca" &&

                            (
                                !searchText ||
                                itemMatchesSearch(
                                    item,
                                    searchText
                                )
                            )
                        );

                    }
                );

        }
        else if (
            activeFilter === "contractor"
        ) {

            byId(
                "nsv-contractor-kpi"
            ).classList.add(
                "nsv-kpi-active"
            );


            filteredItems =
                allItems.filter(
                    function (item) {

                        var organisation =
                            getFieldValue(
                                item,
                                "field_2"
                            )
                                .toLowerCase();


                        return (
                            organisation.includes(
                                "contractor"
                            ) &&

                            (
                                !searchText ||
                                itemMatchesSearch(
                                    item,
                                    searchText
                                )
                            )
                        );

                    }
                );

        }
        else if (
            activeFilter === "expired"
        ) {

            filteredItems =
                allItems.filter(
                    function (item) {

                        return (
                            isExpired(
                                item
                            ) &&

                            (
                                !searchText ||
                                itemMatchesSearch(
                                    item,
                                    searchText
                                )
                            )
                        );

                    }
                );

        }
        else if (
            activeFilter === "renewal"
        ) {

            filteredItems =
                allItems.filter(
                    function (item) {

                        return (
                            isDueForRenewal(
                                item
                            ) &&

                            (
                                !searchText ||
                                itemMatchesSearch(
                                    item,
                                    searchText
                                )
                            )
                        );

                    }
                );

        }
        else {

            filteredItems =
                allItems.filter(
                    function (item) {

                        return (
                            !searchText ||
                            itemMatchesSearch(
                                item,
                                searchText
                            )
                        );

                    }
                );

        }


        updateExpiredButton();

        updateRenewalButton();

    }


    function applyOrganisationFilter(
        filterType
    ) {

        exitApplicationsModeIfActive();

        activeFilter =
            filterType;


        applyCurrentFilter();

        renderTable();

        updatePanelDescription();

    }


    function applyExpiredFilter() {

        exitApplicationsModeIfActive();

        activeFilter =
            "expired";


        applyCurrentFilter();

        renderTable();

        updatePanelDescription();

    }


    function applyRenewalFilter() {

        exitApplicationsModeIfActive();

        activeFilter =
            "renewal";


        applyCurrentFilter();

        renderTable();

        updatePanelDescription();

    }


    function clearAllFilters() {

        exitApplicationsModeIfActive();

        activeFilter =
            null;


        applyCurrentFilter();

        renderTable();

        updatePanelDescription();

    }


    function toggleMcaFilter() {

        if (
            activeFilter === "mca"
        ) {

            clearAllFilters();

        }
        else {

            applyOrganisationFilter(
                "mca"
            );

        }

    }


    function toggleContractorFilter() {

        if (
            activeFilter === "contractor"
        ) {

            clearAllFilters();

        }
        else {

            applyOrganisationFilter(
                "contractor"
            );

        }

    }


    function toggleExpiredFilter() {

        if (
            activeFilter === "expired"
        ) {

            clearAllFilters();

        }
        else {

            applyExpiredFilter();

        }

    }


    function toggleRenewalFilter() {

        if (
            activeFilter === "renewal"
        ) {

            clearAllFilters();

        }
        else {

            applyRenewalFilter();

        }

    }


    /* ============================================================
       SEARCH
       ============================================================ */

    function search(text) {

        if (applicationsMode) {

            applyApplicationsFilter();

            renderApplicationsTable();

            updatePanelDescription();

            return;

        }


        applyCurrentFilter();

        renderTable();

        updatePanelDescription();

    }


    /* ============================================================
       PANEL DESCRIPTION
       ============================================================ */

    function updatePanelDescription() {

        var description =
            byId(
                "nsv-panel-description"
            );


        if (!description) {

            return;

        }


        if (applicationsMode) {

            var applicationCount =
                filteredApplicationItems.length;


            description.textContent =
                applicationCount.toLocaleString(
                    "en-GB"
                ) +
                " application record" +
                (
                    applicationCount === 1
                        ? ""
                        : "s"
                );


            if (applicationSortState.field) {

                var sortedApplicationColumn =
                    columns.find(
                        function (column) {

                            return (
                                column.field ===
                                applicationSortState.field
                            );

                        }
                    );


                if (sortedApplicationColumn) {

                    description.textContent +=
                        " • Sorted by " +
                        sortedApplicationColumn.header +
                        " (" +
                        (
                            applicationSortState.direction === "asc"
                                ? "A–Z"
                                : "Z–A"
                        ) +
                        ")";

                }

            }


            return;

        }


        var count =
            filteredItems.length;


        if (
            activeFilter === "mca"
        ) {

            description.textContent =
                count.toLocaleString(
                    "en-GB"
                ) +
                " MCA records";

        }
        else if (
            activeFilter === "contractor"
        ) {

            description.textContent =
                count.toLocaleString(
                    "en-GB"
                ) +
                " Contractor records";

        }
        else if (
            activeFilter === "expired"
        ) {

            description.textContent =
                count.toLocaleString(
                    "en-GB"
                ) +
                " expired records";

        }
        else if (
            activeFilter === "renewal"
        ) {

            description.textContent =
                count.toLocaleString(
                    "en-GB"
                ) +
                " records due for renewal within 90 days";

        }
        else {

            description.textContent =
                count.toLocaleString(
                    "en-GB"
                ) +
                " SharePoint list records";

        }


        if (sortState.field) {

            var sortedColumn =
                columns.find(
                    function (column) {

                        return (
                            column.field ===
                            sortState.field
                        );

                    }
                );


            if (sortedColumn) {

                description.textContent +=
                    " • Sorted by " +
                    sortedColumn.header +
                    " (" +
                    (
                        sortState.direction === "asc"
                            ? "A–Z"
                            : "Z–A"
                    ) +
                    ")";

            }

        }

    }


    /* ============================================================
       SORTING
       ============================================================ */

    function getSortValue(
        item,
        column
    ) {

        var value =
            item[
                column.field
            ];


        if (
            value === null ||
            value === undefined
        ) {

            return "";

        }


        if (
            column.type === "date"
        ) {

            var date =
                new Date(value);


            if (
                !isNaN(
                    date.getTime()
                )
            ) {

                return date.getTime();

            }


            return 0;

        }


        return String(value)
            .trim()
            .toLowerCase();

    }


    function sortItems() {

        if (!sortState.field) {

            return;

        }


        var column =
            columns.find(
                function (item) {

                    return (
                        item.field ===
                        sortState.field
                    );

                }
            );


        if (!column) {

            return;

        }


        filteredItems.sort(
            function (a, b) {

                var valueA =
                    getSortValue(
                        a,
                        column
                    );


                var valueB =
                    getSortValue(
                        b,
                        column
                    );


                if (
                    valueA === valueB
                ) {

                    return 0;

                }


                var result =
                    valueA <
                    valueB
                        ? -1
                        : 1;


                return (
                    sortState.direction === "asc"
                        ? result
                        : -result
                );

            }
        );

    }


    function handleColumnSort(
        field
    ) {

        if (
            sortState.field === field
        ) {

            sortState.direction =
                sortState.direction === "asc"
                    ? "desc"
                    : "asc";

        }
        else {

            sortState.field =
                field;

            sortState.direction =
                "asc";

        }


        sortItems();

        renderTable();

        updatePanelDescription();

    }


function buildLineChartSvg(labels, values) {

    var width = 640;
    var height = 260;
    var paddingLeft = 40;
    var paddingRight = 20;
    var paddingTop = 20;
    var paddingBottom = 40;

    var maxValue =
        Math.max.apply(null, values.concat([1]));

    var chartWidth =
        width - paddingLeft - paddingRight;

    var chartHeight =
        height - paddingTop - paddingBottom;

    var stepX =
        values.length > 1
            ? chartWidth / (values.length - 1)
            : 0;


    var points =
        values.map(function (value, index) {

            var x =
                paddingLeft + (stepX * index);

            var y =
                paddingTop + chartHeight -
                ((value / maxValue) * chartHeight);

            return { x: x, y: y, value: value };

        });


    var pathData =
        points.map(function (point, index) {
            return (
                (index === 0 ? "M" : "L") +
                point.x.toFixed(1) + "," +
                point.y.toFixed(1)
            );
        }).join(" ");


    var circles =
        points.map(function (point) {
            return (
                '<circle cx="' + point.x.toFixed(1) +
                '" cy="' + point.y.toFixed(1) +
                '" r="3" fill="var(--nsv-accent)"></circle>'
            );
        }).join("");


    var labelInterval =
        Math.max(1, Math.ceil(labels.length / 12));

    var labelsHtml =
        points.map(function (point, index) {

            if (
                index % labelInterval !== 0 &&
                index !== points.length - 1
            ) {
                return "";
            }

            return (
                '<text x="' + point.x.toFixed(1) +
                '" y="' + (height - 12) +
                '" font-size="10" text-anchor="middle" fill="var(--nsv-muted)">' +
                escapeHtml(labels[index]) +
                '</text>'
            );

        }).join("");


    var gridLines = "";
    var gridCount = 4;

    for (var i = 0; i <= gridCount; i++) {

        var gridY =
            paddingTop + (chartHeight / gridCount) * i;

        var gridValue =
            Math.round(maxValue - (maxValue / gridCount) * i);

        gridLines +=
            '<line x1="' + paddingLeft +
            '" y1="' + gridY.toFixed(1) +
            '" x2="' + (width - paddingRight) +
            '" y2="' + gridY.toFixed(1) +
            '" stroke="var(--nsv-border)" stroke-width="1"></line>';

        gridLines +=
            '<text x="' + (paddingLeft - 8) +
            '" y="' + (gridY + 4).toFixed(1) +
            '" font-size="10" text-anchor="end" fill="var(--nsv-muted)">' +
            gridValue +
            '</text>';

    }


    return (
        '<svg viewBox="0 0 ' + width + ' ' + height +
        '" class="nsv-chart-svg">' +
        gridLines +
        '<path d="' + pathData +
        '" fill="none" stroke="var(--nsv-accent)" stroke-width="2"></path>' +
        circles +
        labelsHtml +
        '</svg>'
    );

}

    /* ============================================================
       RENDER TABLE
       ============================================================ */

    function renderTable() {

        var tableContainer =
            byId(
                "nsv-table-container"
            );


        if (
            !filteredItems.length
        ) {

            tableContainer.innerHTML = `

                <div class="nsv-empty">
                    No records found.
                </div>

            `;

            return;

        }


        sortItems();


        var html = `

            <table class="nsv-table">

                <colgroup>

        `;


        columns.forEach(
            function (column) {

                var width =
                    columnWidths[
                        column.field
                    ] ||
                    getDefaultColumnWidth(
                        column.field
                    );


                html += `

                    <col
                        data-column-field="${escapeHtml(
                            column.field
                        )}"
                        style="width:${width}px;"
                    >

                `;

            }
        );


        html += `

                </colgroup>

                <thead>

                    <tr>

        `;


        columns.forEach(
            function (column, index) {

                var sortIndicator =
                    "";


                if (
                    sortState.field ===
                    column.field
                ) {

                    sortIndicator =
                        sortState.direction === "asc"
                            ? " ▲"
                            : " ▼";

                }


                html += `

                    <th
                        data-column-index="${index}"
                        data-column-field="${escapeHtml(
                            column.field
                        )}"
                        class="nsv-sortable-header"
                        title="Click to sort"
                    >

                        <div class="nsv-th-content">

                            <span
                                class="nsv-sort-label"
                                data-sort-field="${escapeHtml(
                                    column.field
                                )}"
                            >
                                ${escapeHtml(
                                    column.header
                                )}${sortIndicator}
                            </span>

                            <span
                                class="nsv-column-resizer"
                                data-column-index="${index}"
                                title="Drag to resize column"
                            ></span>

                        </div>

                    </th>

                `;

            }
        );


        html += `

                    </tr>

                </thead>

                <tbody>

        `;


        filteredItems.forEach(
            function (item) {

                var rowClass =
                    updateMode
                        ? "nsv-row-editable"
                        : "nsv-row-clickable";


                html += `

                    <tr
                        class="${rowClass}"
                        data-item-id="${escapeHtml(
                            item.ID
                        )}"
                    >

                `;


                columns.forEach(
                    function (column) {

                        html += `

                            <td
                                data-column-field="${escapeHtml(
                                    column.field
                                )}"
                            >
                                ${formatValue(
                                    item[
                                        column.field
                                    ],
                                    column.type
                                )}
                            </td>

                        `;

                    }
                );


                html += `

                    </tr>

                `;

            }
        );


        html += `

                </tbody>

            </table>

        `;


        tableContainer.innerHTML =
            html;


        attachColumnResizeHandlers();

        attachColumnSortHandlers();

        attachRowClickHandlers();

    };

    function renderDashboard() {

    var tableContainer =
        byId("nsv-table-container");


    if (!tableContainer) {
        return;
    }


    var data =
        dashboardPeriod === 1
            ? getDailyCreatedCounts()
            : getMonthlyCreatedCounts(dashboardPeriod);


    tableContainer.innerHTML = `

        <div class="nsv-dashboard-grid">

            <div class="nsv-dashboard-card nsv-dashboard-card-wide">

                <div class="nsv-dashboard-card-header">

                    <h3 class="nsv-dashboard-card-title">
                        Clearances added
                    </h3>

                    <select
                        id="nsv-dashboard-period"
                        class="nsv-dashboard-select"
                    >
                        <option value="1" ${dashboardPeriod === 1 ? "selected" : ""}>This month</option>
                        <option value="2" ${dashboardPeriod === 2 ? "selected" : ""}>Last 2 months</option>
                        <option value="3" ${dashboardPeriod === 3 ? "selected" : ""}>Last 3 months</option>
                    </select>

                </div>

                <div id="nsv-dashboard-chart" class="nsv-dashboard-chart">
                    ${buildLineChartSvg(data.labels, data.values)}
                </div>

            </div>


            <div class="nsv-dashboard-card">

                <h3 class="nsv-dashboard-card-title">
                    Currently being processed
                </h3>

                <div class="nsv-dashboard-placeholder">
                    Coming soon
                </div>

            </div>


            <div class="nsv-dashboard-card">

                <h3 class="nsv-dashboard-card-title">
                    Not renewed in time
                </h3>

                <div class="nsv-dashboard-placeholder">
                    Coming soon
                </div>

            </div>

        </div>

    `;


    var periodSelect =
        byId("nsv-dashboard-period");


    if (periodSelect) {

        periodSelect.addEventListener(
            "change",
            function () {

                dashboardPeriod =
                    Number(periodSelect.value);

                renderDashboard();

            }
        );

    }

}


function toggleDashboardMode() {

    dashboardMode = !dashboardMode;


    var button =
        byId("nsv-dashboards");

    var applicationsButton =
        byId("nsv-applications");

    var searchInput =
        byId("nsv-search");


    if (dashboardMode) {

        /*
         * Dashboards, Records and Applications are mutually
         * exclusive views. Turning Dashboards on always
         * turns Applications off.
         */
        if (applicationsMode) {

            applicationsMode = false;

            if (applicationsButton) {

                applicationsButton.classList.remove(
                    "nsv-update-active"
                );

                applicationsButton.setAttribute(
                    "aria-pressed",
                    "false"
                );

            }

        }

        button.classList.add("nsv-update-active");
        button.setAttribute("aria-pressed", "true");

        if (searchInput) {
            searchInput.disabled = true;
        }

        renderDashboard();

        setStatus("Viewing dashboards.");

    }
    else {

        button.classList.remove("nsv-update-active");
        button.setAttribute("aria-pressed", "false");

        if (searchInput) {
            searchInput.disabled = false;
        }

        renderTable();

        setStatus("Viewing table.");

    }

}


    /* ============================================================
       DEFAULT COLUMN WIDTHS
       ============================================================ */

    function getDefaultColumnWidth(
        field
    ) {

        var widths = {

            "field_6": 190,

            "field_2": 150,

            "field_5": 170,

            "field_10": 230,

            "field_14": 120,

            "field_13": 145,

            "field_15": 120,

            "field_16": 120,

            "field_9": 170,

            "field_4": 130

        };


        return widths[field] ||
            150;

    }


    /* ============================================================
       COLUMN RESIZING
       ============================================================ */

    function attachColumnResizeHandlers() {

        var resizers =
            document.querySelectorAll(
                ".nsv-column-resizer"
            );


        resizers.forEach(
            function (resizer) {

                resizer.addEventListener(
                    "mousedown",
                    startColumnResize
                );

            }
        );

    }


    var resizeState =
        null;


    function startColumnResize(
        event
    ) {

        event.preventDefault();

        event.stopPropagation();


        var resizer =
            event.currentTarget;


        var columnIndex =
            Number(
                resizer.getAttribute(
                    "data-column-index"
                )
            );


        var table =
            document.querySelector(
                ".nsv-table"
            );


        if (!table) {

            return;

        }


        var headerCells =
            table.querySelectorAll(
                "thead th"
            );


        var headerCell =
            headerCells[
                columnIndex
            ];


        if (!headerCell) {

            return;

        }


        resizeState = {

            columnIndex:
                columnIndex,

            startX:
                event.clientX,

            startWidth:
                headerCell.getBoundingClientRect().width

        };


        document.body.style.cursor =
            "col-resize";

        document.body.style.userSelect =
            "none";


        document.addEventListener(
            "mousemove",
            handleColumnResize
        );

        document.addEventListener(
            "mouseup",
            stopColumnResize
        );

    }


    function handleColumnResize(
        event
    ) {

        if (!resizeState) {

            return;

        }


        var delta =
            event.clientX -
            resizeState.startX;


        var newWidth =
            resizeState.startWidth +
            delta;


        newWidth =
            Math.max(
                90,
                Math.min(
                    600,
                    newWidth
                )
            );


        var column =
            columns[
                resizeState.columnIndex
            ];


        columnWidths[
            column.field
        ] =
            Math.round(
                newWidth
            );


        var table =
            document.querySelector(
                ".nsv-table"
            );


        if (!table) {

            return;

        }


        var col =
            table.querySelector(
                "col:nth-child(" +
                (
                    resizeState.columnIndex +
                    1
                ) +
                ")"
            );


        if (col) {

            col.style.width =
                newWidth +
                "px";

        }


        var headerCells =
            table.querySelectorAll(
                "thead th"
            );


        if (
            headerCells[
                resizeState.columnIndex
            ]
        ) {

            headerCells[
                resizeState.columnIndex
            ].style.width =
                newWidth +
                "px";

        }

    }


    function stopColumnResize() {

        resizeState =
            null;


        document.body.style.cursor =
            "";

        document.body.style.userSelect =
            "";


        document.removeEventListener(
            "mousemove",
            handleColumnResize
        );

        document.removeEventListener(
            "mouseup",
            stopColumnResize
        );

    }


    /* ============================================================
       COLUMN SORT HANDLERS
       ============================================================ */

    function attachColumnSortHandlers() {

        var headers =
            document.querySelectorAll(
                ".nsv-sortable-header"
            );


        headers.forEach(
            function (header) {

                header.addEventListener(
                    "click",
                    function (event) {

                        if (
                            event.target.classList.contains(
                                "nsv-column-resizer"
                            )
                        ) {

                            return;

                        }


                        var field =
                            header.getAttribute(
                                "data-column-field"
                            );


                        if (field) {

                            handleColumnSort(
                                field
                            );

                        }

                    }
                );

            }
        );

    }


    /* ============================================================
       ROW CLICK
       ============================================================ */

    function attachRowClickHandlers() {

        var rows =
            document.querySelectorAll(
                ".nsv-table tbody tr"
            );


        rows.forEach(
            function (row) {

                row.addEventListener(
                    "click",
                    function (event) {

                        if (
                            event.target.classList.contains(
                                "nsv-column-resizer"
                            )
                        ) {

                            return;

                        }


                        var itemId =
                            Number(
                                row.getAttribute(
                                    "data-item-id"
                                )
                            );


                        var item =
                            allItems.find(
                                function (record) {

                                    return Number(
                                        record.ID
                                    ) ===
                                    itemId;

                                }
                            );


                        if (!item) {

                            return;

                        }


                        selectedItem =
                            item;


                        if (updateMode) {

                            openEditModal(
                                item
                            );

                        }
                        else {

                            openReadOnlyModal(
                                item
                            );

                        }

                    }
                );

            }
        );

    }


    /* ============================================================
       UPDATE MODE
       ============================================================ */

    function toggleUpdateMode() {

        exitApplicationsModeIfActive();

        updateMode =
            !updateMode;


        var button =
            byId(
                "nsv-update-clearance"
            );


        var hint =
            byId(
                "nsv-update-hint"
            );


        if (updateMode) {

            button.classList.add(
                "nsv-update-active"
            );


            button.setAttribute(
                "aria-pressed",
                "true"
            );


            hint.style.display =
                "block";


            setStatus(
                "Update mode active. Click a row to edit."
            );

        }
        else {

            button.classList.remove(
                "nsv-update-active"
            );


            button.setAttribute(
                "aria-pressed",
                "false"
            );


            hint.style.display =
                "none";


            setStatus(
                "Update mode disabled."
            );

        }


        renderTable();

    }


    /* ============================================================
       MODAL
       ============================================================ */

async function openModal(
    mode,
    item
) {

        closeModal();


        /*
         * IMPORTANT FIX:
         *
         * Both "add" and "edit" are form modes.
         * Previously only "edit" was treated as a form,
         * which caused "add" to incorrectly open the
         * read-only modal.
         */

        var isFormMode =
            mode === "add" ||
            mode === "edit";


        var isEdit =
            mode === "edit";


        var isAdd =
            mode === "add";


        /*
         * Add mode has no existing item, so use an empty
         * object. This prevents item[column.field] from
         * throwing an exception.
         */

        var formItem =
            item || {};


        var title =
            isAdd
                ? "Add clearance holder"
                : isEdit
                    ? "Update clearance"
                    : "Clearance holder";


        var description =
            isAdd
                ? "Enter the clearance holder's details and optionally upload a clearance PDF."
                : isEdit
                    ? "Update the clearance holder's details and optionally upload a new clearance PDF."
                    : "View the clearance holder's details.";


        var fieldsHtml =
            "";


        TABLE_COLUMNS.forEach(
            function (column) {

                var value =
                    isFormMode
                        ? getFormValue(
                            formItem[
                                column.field
                            ],
                            column.type
                        )
                        : formatModalValue(
                            formItem[
                                column.field
                            ],
                            column.type
                        );


                fieldsHtml +=
                    isFormMode
                        ? buildFormField(
                            column,
                            value
                        )
                        : buildReadOnlyField(
                            column,
                            value
                        );

            }
        );


        var modal =
            document.createElement(
                "div"
            );


        modal.id =
            "nsv-modal-overlay";


        modal.className =
            "nsv-modal-overlay";


        modal.innerHTML = `

            <div
                class="nsv-modal"
                role="dialog"
                aria-modal="true"
                aria-labelledby="nsv-modal-title"
            >

                <div class="nsv-modal-header">

                    <div>

                        <h2
                            id="nsv-modal-title"
                            class="nsv-modal-title"
                        >
                            ${escapeHtml(
                                title
                            )}
                        </h2>

                        <p class="nsv-modal-description">
                            ${escapeHtml(
                                description
                            )}
                        </p>

                    </div>


                    <button
                        type="button"
                        id="nsv-modal-close"
                        class="nsv-modal-close"
                        aria-label="Close"
                    >
                        ×
                    </button>

                </div>


                ${
                    isFormMode
                        ? `
                            <div
                                id="nsv-modal-error"
                                class="nsv-modal-error"
                                style="display:none;"
                            ></div>
                        `
                        : ""
                }


                ${
                    isFormMode
                        ? `
                            <form
                                id="nsv-clearance-form"
                                class="nsv-form"
                            >

                                <div class="nsv-form-grid">

                                    ${fieldsHtml}

                                    <div class="nsv-form-field nsv-upload-field">

                                        <label
                                            class="nsv-form-label"
                                            for="nsv-clearance-file"
                                        >
                                            Upload clearance
                                        </label>

                                        <input
                                            id="nsv-clearance-file"
                                            class="nsv-form-input nsv-file-input"
                                            type="file"
                                            accept=".pdf,application/pdf"
                                        >

                                        <div class="nsv-upload-help">
                                            PDF only. Maximum file size: 25 MB.
                                        </div>

                                        <div
                                            id="nsv-selected-file"
                                            class="nsv-selected-file"
                                        >
                                            No new PDF selected.
                                        </div>

                                        <div class="nsv-upload-existing">
                                            Selecting a PDF will add it to
                                            this item's SharePoint attachments.
                                            Existing attachments will not be removed.
                                        </div>

                                    </div>

                                </div>


                                <div class="nsv-modal-footer">

                                    <button
                                        type="button"
                                        id="nsv-modal-cancel"
                                        class="nsv-modal-button nsv-secondary-button"
                                    >
                                        Cancel
                                    </button>

                                    <button
                                        type="submit"
                                        id="nsv-modal-save"
                                        class="nsv-modal-button nsv-primary-button"
                                    >
                                        ${
                                            isAdd
                                                ? "Add clearance"
                                                : "Save changes"
                                        }
                                    </button>

                                </div>

                            </form>
                        `
                        : `
                            <div class="nsv-form">

                                <div class="nsv-form-grid">

                                    ${fieldsHtml}

                                </div>


                                <div class="nsv-modal-footer">

                                    <button
                                        type="button"
                                        id="nsv-modal-delete"
                                        class="nsv-modal-button nsv-danger-button"
                                    >
                                        Delete record
                                    </button>

                                    <button
                                        type="button"
                                        id="nsv-modal-cancel"
                                        class="nsv-modal-button nsv-secondary-button"
                                    >
                                        Close
                                    </button>

                                </div>

                            </div>
                        `
                }

            </div>

        `;


        document.body.appendChild(
            modal
        );


document.body.appendChild(
    modal
);


/*
 * Existing SharePoint attachment.
 *
 * Only show this for existing records.
 * Add mode has no SharePoint item yet.
 */
if (
    !isAdd &&
    item &&
    item.ID
) {

    var existingAttachment =
        await getExistingAttachment(
            item.ID
        );


    if (existingAttachment) {

        var attachmentField =
            document.createElement(
                "div"
            );


        attachmentField.className =
            "nsv-readonly-field nsv-attachment-field";


attachmentField.innerHTML = `
    <div class="nsv-readonly-label">Attachment</div>
    <div class="nsv-readonly-value">
        <button
            type="button"
            class="nsv-existing-attachment-link nsv-button"
            data-attachment-url="${escapeHtml(existingAttachment.url)}"
        >
            ${escapeHtml(existingAttachment.fileName)}
        </button>
    </div>
`;

var attachmentLink = attachmentField.querySelector(".nsv-existing-attachment-link");

if (attachmentLink) {
    attachmentLink.addEventListener("click", function (event) {
        event.preventDefault();
        event.stopPropagation();

        var attachmentUrl = attachmentLink.getAttribute("data-attachment-url");
        if (!attachmentUrl) { return; }

        window.open(attachmentUrl, "_blank", "noopener,noreferrer");
    });
}

        var formGrid =
            modal.querySelector(
                ".nsv-form-grid"
            );


        if (formGrid) {

            /*
             * Put the attachment at the bottom of
             * the existing modal fields, immediately
             * before the upload control in edit mode.
             */
            var uploadField =
                formGrid.querySelector(
                    ".nsv-upload-field"
                );


            if (
                uploadField &&
                isFormMode
            ) {

                formGrid.insertBefore(
                    attachmentField,
                    uploadField
                );

            }
            else {

                /*
                 * Read-only modal.
                 */
                formGrid.appendChild(
                    attachmentField
                );

            }

        }

    }

}


document.body.classList.add(
    "nsv-modal-open"
);


        /*
         * IMPORTANT FIX:
         *
         * Attach form events for BOTH add and edit.
         */

        if (isFormMode) {

            attachModalEvents(
                mode,
                item
            );


            var fileInput =
                byId(
                    "nsv-clearance-file"
                );


            if (fileInput) {

                fileInput.addEventListener(
                    "change",
                    function () {

                        updateSelectedFileDisplay(
                            fileInput
                        );

                    }
                );

            }


            var firstInput =
                modal.querySelector(
                    ".nsv-form-input"
                );


            if (firstInput) {

                setTimeout(
                    function () {

                        firstInput.focus();

                    },
                    50
                );

            }

        }
        else {

            attachReadOnlyModalEvents();

        }

    }


    function openAddModal() {

        exitApplicationsModeIfActive();

        selectedItem =
            null;


        openModal(
            "add",
            null
        );

    }


    function openEditModal(
        item
    ) {

        selectedItem =
            item;


        openModal(
            "edit",
            item
        );

    }


    function openReadOnlyModal(
        item
    ) {

        selectedItem =
            item;


        openModal(
            "view",
            item
        );

    }


    function closeModal() {

        var modal =
            byId(
                "nsv-modal-overlay"
            );


        if (modal) {

            modal.remove();

        }


        document.body.classList.remove(
            "nsv-modal-open"
        );


        document.removeEventListener(
            "keydown",
            modalEscapeHandler
        );

    }


    /* ============================================================
       READ-ONLY MODAL FIELD
       ============================================================ */

    function buildReadOnlyField(
        column,
        value
    ) {

        return `

            <div class="nsv-form-field">

                <label
                    class="nsv-form-label"
                >
                    ${escapeHtml(
                        column.header
                    )}
                </label>

                <div
                    class="nsv-form-input nsv-readonly-value"
                    aria-readonly="true"
                >
                    ${
                        value ||
                        "—"
                    }
                </div>

            </div>

        `;

    }


    function formatModalValue(
        value,
        type
    ) {

        if (
            value === null ||
            value === undefined ||
            value === ""
        ) {

            return "—";

        }


        if (
            type === "date"
        ) {

            return escapeHtml(
                formatDate(
                    value
                )
            );

        }


        if (
            typeof value === "object"
        ) {

            if (value.Title) {

                return escapeHtml(
                    value.Title
                );

            }


            if (value.Name) {

                return escapeHtml(
                    value.Name
                );

            }


            try {

                return escapeHtml(
                    JSON.stringify(
                        value
                    )
                );

            }
            catch (e) {

                return "[Object]";

            }

        }


        return escapeHtml(
            String(value)
        );

    }


    function attachReadOnlyModalEvents() {

        var closeButton =
            byId(
                "nsv-modal-close"
            );


        var cancelButton =
            byId(
                "nsv-modal-cancel"
            );


        var deleteButton =
            byId(
                "nsv-modal-delete"
            );


        if (closeButton) {

            closeButton.addEventListener(
                "click",
                closeModal
            );

        }


        if (cancelButton) {

            cancelButton.addEventListener(
                "click",
                closeModal
            );

        }


        if (deleteButton) {

            deleteButton.addEventListener(
                "click",
                function () {

                    if (selectedItem) {

                        openDeleteConfirmation(
                            selectedItem
                        );

                    }

                }
            );

        }


        var overlay =
            byId(
                "nsv-modal-overlay"
            );


        if (overlay) {

            overlay.addEventListener(
                "click",
                function (event) {

                    if (
                        event.target ===
                        overlay
                    ) {

                        closeModal();

                    }

                }
            );

        }


        document.addEventListener(
            "keydown",
            modalEscapeHandler
        );

    }


    /* ============================================================
       DELETE CONFIRMATION
       ============================================================ */

    function openDeleteConfirmation(item) {

        var existing =
            byId(
                "nsv-delete-confirmation"
            );


        if (existing) {

            existing.remove();

        }


        var confirmation =
            document.createElement(
                "div"
            );


        confirmation.id =
            "nsv-delete-confirmation";


        confirmation.className =
            "nsv-modal-overlay";


        confirmation.style.zIndex =
            "10001";


        confirmation.innerHTML = `

            <div
                class="nsv-modal nsv-confirmation-modal"
                role="dialog"
                aria-modal="true"
                aria-labelledby="nsv-delete-confirmation-title"
            >

                <div class="nsv-modal-header">

                    <div>

                        <h2
                            id="nsv-delete-confirmation-title"
                            class="nsv-modal-title"
                        >
                            Confirm deletion
                        </h2>

                        <p class="nsv-modal-description">
                            Are you sure you wish to permanently delete this record?
                        </p>

                    </div>

                    <button
                        type="button"
                        id="nsv-delete-confirmation-close"
                        class="nsv-modal-close"
                        aria-label="Close"
                    >
                        ×
                    </button>

                </div>

                <div class="nsv-modal-footer">

                    <button
                        type="button"
                        id="nsv-delete-no"
                        class="nsv-modal-button nsv-secondary-button"
                    >
                        No
                    </button>

                    <button
                        type="button"
                        id="nsv-delete-yes"
                        class="nsv-modal-button nsv-danger-button"
                    >
                        Yes, permanently delete
                    </button>

                </div>

            </div>

        `;


        document.body.appendChild(
            confirmation
        );


        byId(
            "nsv-delete-confirmation-close"
        ).addEventListener(
            "click",
            closeDeleteConfirmation
        );


        byId(
            "nsv-delete-no"
        ).addEventListener(
            "click",
            closeDeleteConfirmation
        );


        byId(
            "nsv-delete-yes"
        ).addEventListener(
            "click",
            function () {

                permanentlyDeleteItem(
                    item
                );

            }
        );


        confirmation.addEventListener(
            "keydown",
            function (event) {

                if (event.key === "Escape") {

                    event.preventDefault();

                    event.stopPropagation();

                    closeDeleteConfirmation();

                }

            }
        );


        confirmation.addEventListener(
            "click",
            function (event) {

                if (
                    event.target ===
                    confirmation
                ) {

                    closeDeleteConfirmation();

                }

            }
        );


        setTimeout(
            function () {

                var yesButton =
                    byId(
                        "nsv-delete-yes"
                    );


                if (yesButton) {

                    yesButton.focus();

                }

            },
            50
        );

    }


    function closeDeleteConfirmation() {

        var confirmation =
            byId(
                "nsv-delete-confirmation"
            );


        if (confirmation) {

            confirmation.remove();

        }

    }


    async function permanentlyDeleteItem(
        item
    ) {

        if (
            !item ||
            !item.ID
        ) {

            closeDeleteConfirmation();

            return;

        }


        var yesButton =
            byId(
                "nsv-delete-yes"
            );


        var noButton =
            byId(
                "nsv-delete-no"
            );


        if (yesButton) {

            yesButton.disabled =
                true;

            yesButton.textContent =
                "Deleting...";

        }


        if (noButton) {

            noButton.disabled =
                true;

        }


        try {

            setStatus(
                "Permanently deleting clearance record..."
            );


            var itemUrl =
                SITE_URL +
                "/_api/web/lists/GetByTitle('" +
                encodeURIComponent(
                    LIST_TITLE
                ) +
                "')/items(" +
                item.ID +
                ")";


            var digest =
                await getRequestDigest();


            await apiDelete(
                itemUrl,
                {
                    "X-RequestDigest":
                        digest
                }
            );


            closeDeleteConfirmation();

            closeModal();

            selectedItem =
                null;


            setStatus(
                "Clearance record deleted successfully."
            );


            await initialise();

        }
        catch (error) {

            console.error(
                "Delete clearance error:",
                error
            );


            if (yesButton) {

                yesButton.disabled =
                    false;

                yesButton.textContent =
                    "Yes, permanently delete";

            }


            if (noButton) {

                noButton.disabled =
                    false;

            }


            var message =
                error &&
                error.message
                    ? error.message
                    : "Unable to delete the record.";


            setStatus(
                "ERROR: " +
                message
            );


            alert(
                "The record could not be deleted.\n\n" +
                message
            );

        }

    }


    /* ============================================================
       FORM FIELD GENERATION
       ============================================================ */

    function buildFormField(
        column,
        value
    ) {

        var inputType =
            column.type === "date"
                ? "date"
                : "text";


        return `

            <div class="nsv-form-field">

                <label
                    class="nsv-form-label"
                    for="nsv-field-${escapeHtml(
                        column.field
                    )}"
                >
                    ${escapeHtml(
                        column.header
                    )}
                </label>


                <input
                    id="nsv-field-${escapeHtml(
                        column.field
                    )}"
                    class="nsv-form-input"
                    data-field="${escapeHtml(
                        column.field
                    )}"
                    type="${inputType}"
                    value="${escapeHtml(
                        value
                    )}"
                    ${
                        column.field === "field_10"
                            ? 'autocomplete="email"'
                            : ""
                    }
                >

            </div>

        `;

    }


    function getFormValue(
        value,
        type
    ) {

        if (
            value === null ||
            value === undefined ||
            value === ""
        ) {

            return "";

        }


        if (
            type === "date"
        ) {

            var date =
                new Date(value);


            if (
                !isNaN(
                    date.getTime()
                )
            ) {

                var year =
                    date.getFullYear();


                var month =
                    String(
                        date.getMonth() + 1
                    )
                        .padStart(
                            2,
                            "0"
                        );


                var day =
                    String(
                        date.getDate()
                    )
                        .padStart(
                            2,
                            "0"
                        );


                return (
                    year +
                    "-" +
                    month +
                    "-" +
                    day
                );

            }

        }


        return String(value);

    }


    /* ============================================================
       MODAL EVENTS
       ============================================================ */

    function attachModalEvents(
        mode,
        item
    ) {

        var closeButton =
            byId(
                "nsv-modal-close"
            );


        var cancelButton =
            byId(
                "nsv-modal-cancel"
            );


        var form =
            byId(
                "nsv-clearance-form"
            );


        if (closeButton) {

            closeButton.addEventListener(
                "click",
                closeModal
            );

        }


        if (cancelButton) {

            cancelButton.addEventListener(
                "click",
                closeModal
            );

        }


        if (form) {

            form.addEventListener(
                "submit",
                function (event) {

                    event.preventDefault();


                    if (
                        mode === "edit"
                    ) {

                        updateClearance(
                            item
                        );

                    }
                    else if (
                        mode === "add"
                    ) {

                        addClearance();

                    }

                }
            );

        }


        var overlay =
            byId(
                "nsv-modal-overlay"
            );


        if (overlay) {

            overlay.addEventListener(
                "click",
                function (event) {

                    if (
                        event.target ===
                        overlay
                    ) {

                        closeModal();

                    }

                }
            );

        }


        document.addEventListener(
            "keydown",
            modalEscapeHandler
        );

    }


    function modalEscapeHandler(
        event
    ) {

        if (
            event.key === "Escape"
        ) {

            closeModal();

        }

    }


    /* ============================================================
       FILE DISPLAY
       ============================================================ */

    function updateSelectedFileDisplay(
        input
    ) {

        var display =
            byId(
                "nsv-selected-file"
            );


        if (!display) {

            return;

        }


        if (
            !input.files ||
            !input.files.length
        ) {

            display.textContent =
                "No new PDF selected.";


            display.classList.remove(
                "nsv-file-selected"
            );


            display.classList.remove(
                "nsv-file-invalid"
            );


            return;

        }


        var file =
            input.files[0];


        var extension =
            file.name
                .split(".")
                .pop()
                .toLowerCase();


        if (
            extension !== "pdf"
        ) {

            display.textContent =
                "Invalid file. Please select a PDF.";


            display.classList.remove(
                "nsv-file-selected"
            );


            display.classList.add(
                "nsv-file-invalid"
            );


            return;

        }


        if (
            file.size >
            MAX_ATTACHMENT_SIZE
        ) {

            display.textContent =
                "File is too large. Maximum size is 25 MB.";


            display.classList.remove(
                "nsv-file-selected"
            );


            display.classList.add(
                "nsv-file-invalid"
            );


            return;

        }


        display.classList.remove(
            "nsv-file-invalid"
        );


        display.classList.add(
            "nsv-file-selected"
        );


        display.textContent =
            file.name +
            " (" +
            formatFileSize(
                file.size
            ) +
            ")";

    }


    function formatFileSize(
        bytes
    ) {

        if (
            bytes <
            1024
        ) {

            return bytes +
                " B";

        }


        if (
            bytes <
            1024 * 1024
        ) {

            return (
                bytes /
                1024
            ).toFixed(1) +
            " KB";

        }


        return (
            bytes /
            (
                1024 *
                1024
            )
        ).toFixed(1) +
        " MB";

    }


    /* ============================================================
       FORM DATA
       ============================================================ */

    function collectFormData() {

        var data = {};


        TABLE_COLUMNS.forEach(
            function (column) {

                var input =
                    document.querySelector(
                        '[data-field="' +
                        column.field +
                        '"]'
                    );


                if (!input) {

                    return;

                }


                var value =
                    input.value;


                if (
                    column.type === "date"
                ) {

                    if (value) {

                        value =
                            value +
                            "T00:00:00Z";

                    }
                    else {

                        value =
                            null;

                    }

                }


                data[
                    column.field
                ] =
                    value;

            }
        );


        return data;

    }


    function getSelectedPdf() {

        var input =
            byId(
                "nsv-clearance-file"
            );


        if (
            !input ||
            !input.files ||
            !input.files.length
        ) {

            return null;

        }


        return input.files[0];

    }


    /* ============================================================
       VALIDATION
       ============================================================ */

    function validateForm(
        data,
        file
    ) {

        if (
            !String(
                data.field_6 || ""
            ).trim()
        ) {

            return (
                "Please enter the Full Name."
            );

        }


        if (
            data.field_10 &&
            !isValidEmail(
                data.field_10
            )
        ) {

            return (
                "Please enter a valid Email Address."
            );

        }


        if (file) {

            var extension =
                String(
                    file.name || ""
                )
                    .split(".")
                    .pop()
                    .toLowerCase();


            if (
                extension !== "pdf"
            ) {

                return (
                    "The clearance document must be a PDF."
                );

            }


            if (
                file.size >
                MAX_ATTACHMENT_SIZE
            ) {

                return (
                    "The PDF is too large. The maximum permitted size is 25 MB."
                );

            }

        }


        return null;

    }


    function isValidEmail(
        email
    ) {

        return /^[^\s@]+@[^\s@]+\.[^\s@]+$/
            .test(
                String(email).trim()
            );

    }


    /* ============================================================
       ADD CLEARANCE
       ============================================================ */

    async function addClearance() {

        var saveButton =
            byId(
                "nsv-modal-save"
            );


        var data =
            collectFormData();


        var file =
            getSelectedPdf();


        var validationError =
            validateForm(
                data,
                file
            );


        if (
            validationError
        ) {

            showModalError(
                validationError
            );

            return;

        }


        if (!saveButton) {

            console.error(
                "Add clearance: save button was not found."
            );

            return;

        }


        try {

            saveButton.disabled =
                true;


            saveButton.classList.add(
                "nsv-button-loading"
            );


            saveButton.textContent =
                "Adding...";


            setStatus(
                "Creating new clearance holder..."
            );


            var itemUrl =

                SITE_URL +

                "/_api/web/lists/GetByTitle('" +

                encodeURIComponent(
                    LIST_TITLE
                ) +

                "')/items";


            var digest =
                await getRequestDigest();


            var createdItem =
                await apiPost(
                    itemUrl,
                    JSON.stringify(
                        data
                    ),
                    {
                        "Content-Type":
                            "application/json;odata=nometadata",

                        "X-RequestDigest":
                            digest
                    }
                );


            var newItemId =
                null;


            if (
                createdItem &&
                createdItem.ID
            ) {

                newItemId =
                    createdItem.ID;

            }
            else if (
                createdItem &&
                createdItem.Id
            ) {

                newItemId =
                    createdItem.Id;

            }
            else if (
                createdItem &&
                createdItem.d &&
                createdItem.d.Id
            ) {

                newItemId =
                    createdItem.d.Id;

            }
            else if (
                createdItem &&
                createdItem.d &&
                createdItem.d.ID
            ) {

                newItemId =
                    createdItem.d.ID;

            }


            if (!newItemId) {

                var lookupUrl =

                    SITE_URL +

                    "/_api/web/lists/GetByTitle('" +

                    encodeURIComponent(
                        LIST_TITLE
                    ) +

                    "')/items?$select=Id,ID,field_6,Created&$orderby=Created%20desc&$top=10";


                var lookupResult =
                    await apiGet(
                        lookupUrl
                    );


                var matchingItem =
                    (
                        lookupResult.value ||
                        []
                    ).find(
                        function (record) {

                            return (
                                String(
                                    record.field_6 ||
                                    ""
                                ).trim()
                                ===
                                String(
                                    data.field_6 ||
                                    ""
                                ).trim()
                            );

                        }
                    );


                if (
                    matchingItem
                ) {

                    newItemId =
                        matchingItem.ID ||
                        matchingItem.Id;

                }

            }


            if (!newItemId) {

                throw new Error(
                    "The clearance holder was created, but SharePoint did not return the new item's ID. The PDF could not be attached."
                );

            }


            if (file) {

                saveButton.textContent =
                    "Uploading PDF...";


                await uploadPdfAttachment(
                    newItemId,
                    file
                );

            }


            setStatus(
                file
                    ? "Clearance holder and PDF added successfully."
                    : "Clearance holder added successfully."
            );


            closeModal();


            await initialise();

        }
        catch (error) {

            console.error(
                "Add clearance error:",
                error
            );


            showModalError(
                error && error.message
                    ? error.message
                    : "Unable to add the clearance holder."
            );


            saveButton.disabled =
                false;


            saveButton.classList.remove(
                "nsv-button-loading"
            );


            saveButton.textContent =
                "Add clearance";

        }

    }


    /* ============================================================
       UPDATE CLEARANCE
       ============================================================ */

    async function updateClearance(
        item
    ) {

        var saveButton =
            byId(
                "nsv-modal-save"
            );


        var data =
            collectFormData();


        var file =
            getSelectedPdf();


        var validationError =
            validateForm(
                data,
                file
            );


        if (
            validationError
        ) {

            showModalError(
                validationError
            );

            return;

        }


        if (
            !item ||
            !item.ID
        ) {

            showModalError(
                "The selected SharePoint item could not be identified."
            );

            return;

        }


        if (!saveButton) {

            showModalError(
                "The save button could not be found."
            );

            return;

        }


        try {

            saveButton.disabled =
                true;


            saveButton.classList.add(
                "nsv-button-loading"
            );


            saveButton.textContent =
                "Saving...";


            setStatus(
                "Updating clearance holder..."
            );


            var itemUrl =

                SITE_URL +

                "/_api/web/lists/GetByTitle('" +

                encodeURIComponent(
                    LIST_TITLE
                ) +

                "')/items(" +

                item.ID +

                ")";


            var digest =
                await getRequestDigest();


            await apiMerge(
                itemUrl,
                data,
                {
                    "X-RequestDigest":
                        digest
                }
            );


            if (file) {

                saveButton.textContent =
                    "Uploading PDF...";


                await uploadPdfAttachment(
                    item.ID,
                    file
                );

            }


            setStatus(
                file
                    ? "Clearance updated and PDF uploaded successfully."
                    : "Clearance updated successfully."
            );


            closeModal();


            await initialise();

        }
        catch (error) {

            console.error(
                "Update clearance error:",
                error
            );


            showModalError(
                error && error.message
                    ? error.message
                    : "Unable to update the clearance."
            );


            saveButton.disabled =
                false;


            saveButton.classList.remove(
                "nsv-button-loading"
            );


            saveButton.textContent =
                "Save changes";

        }

    }


    /* ============================================================
       MODAL ERROR
       ============================================================ */

    function showModalError(
        message
    ) {

        var element =
            byId(
                "nsv-modal-error"
            );


        if (!element) {

            return;

        }


        element.textContent =
            message;


        element.style.display =
            "block";


        element.scrollIntoView({
            behavior: "smooth",
            block: "nearest"
        });

    }

/* ============================================================
   EXISTING SHAREPOINT ATTACHMENT
   ============================================================ */

async function getExistingAttachment(itemId) {

    if (!itemId) {

        return null;

    }


    var url =

        SITE_URL +

        "/_api/web/lists/GetByTitle('" +

        encodeURIComponent(
            LIST_TITLE
        ) +

        "')/items(" +

        itemId +

        ")/AttachmentFiles?$select=FileName,ServerRelativeUrl";


    try {

        var result =
            await apiGet(
                url
            );


        var attachments =
            result.value ||
            (
                result.d &&
                result.d.results
                    ? result.d.results
                    : []
            );


        if (
            !attachments ||
            !attachments.length
        ) {

            return null;

        }


        var attachment =
            attachments[0];


        if (
            !attachment.FileName ||
            !attachment.ServerRelativeUrl
        ) {

            return null;

        }


        return {

            fileName:
                attachment.FileName,

            url:
                attachment.ServerRelativeUrl.indexOf(
                    "http://"
                ) === 0 ||
                attachment.ServerRelativeUrl.indexOf(
                    "https://"
                ) === 0

                    ? attachment.ServerRelativeUrl

                    : window.location.origin +
                      attachment.ServerRelativeUrl

        };

    }
    catch (error) {

        console.error(
            "Unable to load existing SharePoint attachment:",
            error
        );


        return null;

    }

}

    /* ============================================================
       VALUE FORMATTING
       ============================================================ */

    function formatValue(
        value,
        type
    ) {

        if (
            value === null ||
            value === undefined ||
            value === ""
        ) {

            return "—";

        }


        if (
            type === "date"
        ) {

            return escapeHtml(
                formatDate(
                    value
                )
            );

        }


        if (
            typeof value === "object"
        ) {

            if (
                value.Title
            ) {

                return escapeHtml(
                    value.Title
                );

            }


            if (
                value.Name
            ) {

                return escapeHtml(
                    value.Name
                );

            }


            try {

                return escapeHtml(
                    JSON.stringify(
                        value
                    )
                );

            }
            catch (e) {

                return "[Object]";

            }

        }


        return escapeHtml(
            String(value)
        );

    }


    /* ============================================================
       DATE FORMATTING
       ============================================================ */

    function formatDate(
        value
    ) {

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

            return "—";

        }


        return date.toLocaleDateString(
            "en-GB",
            {
                day:
                    "2-digit",

                month:
                    "short",

                year:
                    "numeric"
            }
        );

    }


    /* ============================================================
       HTML ESCAPING
       ============================================================ */

    function escapeHtml(
        value
    ) {

        return String(value)

            .replace(
                /&/g,
                "&amp;"
            )

            .replace(
                /</g,
                "&lt;"
            )

            .replace(
                />/g,
                "&gt;"
            )

            .replace(
                /"/g,
                "&quot;"
            )

            .replace(
                /'/g,
                "&#039;"
            );

    }


    /* ============================================================
       EVENTS
       ============================================================ */

byId("nsv-dashboards").addEventListener(
    "click",
    function () {
        toggleDashboardMode();
    }
);

byId("nsv-applications").addEventListener(
    "click",
    function () {
        toggleApplicationsMode();
    }
);

    byId(
        "nsv-search"
    ).addEventListener(
        "input",
        function (event) {

            search(
                event.target.value
            );

        }
    );


    byId(
        "nsv-refresh"
    ).addEventListener(
        "click",
        function () {

            initialise();

        }
    );


    byId(
        "nsv-add-clearance"
    ).addEventListener(
        "click",
        function () {

            openAddModal();

        }
    );


    byId(
        "nsv-update-clearance"
    ).addEventListener(
        "click",
        function () {

            toggleUpdateMode();

        }
    );


    byId(
        "nsv-expired-clearance"
    ).addEventListener(
        "click",
        function () {

            toggleExpiredFilter();

        }
    );


    byId(
        "nsv-expired-clearance"
    ).addEventListener(
        "keydown",
        function (event) {

            if (
                event.key === "Enter" ||
                event.key === " "
            ) {

                event.preventDefault();

                toggleExpiredFilter();

            }

        }
    );


    byId(
        "nsv-renewal-filter"
    ).addEventListener(
        "click",
        function () {

            toggleRenewalFilter();

        }
    );


    byId(
        "nsv-renewal-filter"
    ).addEventListener(
        "keydown",
        function (event) {

            if (
                event.key === "Enter" ||
                event.key === " "
            ) {

                event.preventDefault();

                toggleRenewalFilter();

            }

        }
    );


    byId(
        "nsv-mca-kpi"
    ).addEventListener(
        "click",
        function () {

            toggleMcaFilter();

        }
    );


    byId(
        "nsv-mca-kpi"
    ).addEventListener(
        "keydown",
        function (event) {

            if (
                event.key === "Enter" ||
                event.key === " "
            ) {

                event.preventDefault();

                toggleMcaFilter();

            }

        }
    );


    byId(
        "nsv-contractor-kpi"
    ).addEventListener(
        "click",
        function () {

            toggleContractorFilter();

        }
    );


    byId(
        "nsv-contractor-kpi"
    ).addEventListener(
        "keydown",
        function (event) {

            if (
                event.key === "Enter" ||
                event.key === " "
            ) {

                event.preventDefault();

                toggleContractorFilter();

            }

        }
    );


    /* ============================================================
       INITIALISE
       ============================================================ */

    async function initialise() {

        try {

            setStatus(
                "Connecting to SharePoint..."
            );


            await loadList();

            await loadItems();

        }
        catch (error) {

            console.error(
                "NSV Dashboard error:",
                error
            );


            setStatus(
                "ERROR: " +
                error.message
            );


            var tableContainer =
                byId(
                    "nsv-table-container"
                );


            if (tableContainer) {

                tableContainer.innerHTML = `

                    <div class="nsv-error">

                        <h2>
                            Unable to load NSV Workbook
                        </h2>

                        <pre>${escapeHtml(
                            error.message
                        )}</pre>

                    </div>

                `;

            }

        }

    }


    /* ============================================================
       APPLICATIONS: LOAD ITEMS
       ============================================================ */

    async function loadApplicationItems() {

        setStatus(
            "Loading NSV Applications..."
        );


        var url =

            SITE_URL +

            "/_api/web/lists/GetByTitle('" +

            encodeURIComponent(
                APPLICATIONS_LIST_TITLE
            ) +

            "')/items?$top=5000";


        var result =
            await apiGet(
                url
            );


        allApplicationItems =
            result.value || [];


        applyApplicationsFilter();

        renderApplicationsTable();

        updatePanelDescription();


        setStatus(
            allApplicationItems.length +
            " applications loaded successfully."
        );

    }


    /* ============================================================
       APPLICATIONS: FILTER
       ============================================================ */

    function applyApplicationsFilter() {

        var searchInput =
            byId(
                "nsv-search"
            );

        var searchText =
            searchInput
                ? searchInput.value
                    .toLowerCase()
                    .trim()
                : "";


        filteredApplicationItems =
            allApplicationItems.filter(
                function (item) {

                    return (
                        !searchText ||
                        itemMatchesSearch(
                            item,
                            searchText
                        )
                    );

                }
            );

    }


    /* ============================================================
       APPLICATIONS: SORTING
       ============================================================ */

    function sortApplicationItems() {

        if (!applicationSortState.field) {

            return;

        }


        var column =
            columns.find(
                function (item) {

                    return (
                        item.field ===
                        applicationSortState.field
                    );

                }
            );


        if (!column) {

            return;

        }


        filteredApplicationItems.sort(
            function (a, b) {

                var valueA =
                    getSortValue(
                        a,
                        column
                    );


                var valueB =
                    getSortValue(
                        b,
                        column
                    );


                if (
                    valueA === valueB
                ) {

                    return 0;

                }


                var result =
                    valueA <
                    valueB
                        ? -1
                        : 1;


                return (
                    applicationSortState.direction === "asc"
                        ? result
                        : -result
                );

            }
        );

    }


    function handleApplicationColumnSort(
        field
    ) {

        if (
            applicationSortState.field === field
        ) {

            applicationSortState.direction =
                applicationSortState.direction === "asc"
                    ? "desc"
                    : "asc";

        }
        else {

            applicationSortState.field =
                field;

            applicationSortState.direction =
                "asc";

        }


        sortApplicationItems();

        renderApplicationsTable();

        updatePanelDescription();

    }


    /* ============================================================
       APPLICATIONS: RENDER TABLE
       ============================================================ */

    function renderApplicationsTable() {

        var tableContainer =
            byId(
                "nsv-table-container"
            );


        if (
            !filteredApplicationItems.length
        ) {

            tableContainer.innerHTML = `

                <div class="nsv-empty">
                    No applications found.
                </div>

            `;

            return;

        }


        sortApplicationItems();


        var html = `

            <table class="nsv-table">

                <colgroup>

        `;


        columns.forEach(
            function (column) {

                var width =
                    columnWidths[
                        column.field
                    ] ||
                    getDefaultColumnWidth(
                        column.field
                    );


                html += `

                    <col
                        data-column-field="${escapeHtml(
                            column.field
                        )}"
                        style="width:${width}px;"
                    >

                `;

            }
        );


        html += `

                </colgroup>

                <thead>

                    <tr>

        `;


        columns.forEach(
            function (column, index) {

                var sortIndicator =
                    "";


                if (
                    applicationSortState.field ===
                    column.field
                ) {

                    sortIndicator =
                        applicationSortState.direction === "asc"
                            ? " ▲"
                            : " ▼";

                }


                html += `

                    <th
                        data-column-index="${index}"
                        data-column-field="${escapeHtml(
                            column.field
                        )}"
                        class="nsv-sortable-header nsv-applications-sortable-header"
                        title="Click to sort"
                    >

                        <div class="nsv-th-content">

                            <span
                                class="nsv-sort-label"
                                data-sort-field="${escapeHtml(
                                    column.field
                                )}"
                            >
                                ${escapeHtml(
                                    column.header
                                )}${sortIndicator}
                            </span>

                            <span
                                class="nsv-column-resizer"
                                data-column-index="${index}"
                                title="Drag to resize column"
                            ></span>

                        </div>

                    </th>

                `;

            }
        );


        html += `

                    </tr>

                </thead>

                <tbody>

        `;


        filteredApplicationItems.forEach(
            function (item) {

                html += `

                    <tr
                        class="nsv-row-clickable"
                        data-item-id="${escapeHtml(
                            item.ID
                        )}"
                    >

                `;


                columns.forEach(
                    function (column) {

                        html += `

                            <td
                                data-column-field="${escapeHtml(
                                    column.field
                                )}"
                            >
                                ${formatValue(
                                    item[
                                        column.field
                                    ],
                                    column.type
                                )}
                            </td>

                        `;

                    }
                );


                html += `

                    </tr>

                `;

            }
        );


        html += `

                </tbody>

            </table>

        `;


        tableContainer.innerHTML =
            html;


        attachColumnResizeHandlers();

        attachApplicationColumnSortHandlers();

        attachApplicationRowClickHandlers();

    }


    /* ============================================================
       APPLICATIONS: COLUMN SORT HANDLERS
       ============================================================ */

    function attachApplicationColumnSortHandlers() {

        var headers =
            document.querySelectorAll(
                ".nsv-applications-sortable-header"
            );


        headers.forEach(
            function (header) {

                header.addEventListener(
                    "click",
                    function (event) {

                        if (
                            event.target.classList.contains(
                                "nsv-column-resizer"
                            )
                        ) {

                            return;

                        }


                        var field =
                            header.getAttribute(
                                "data-column-field"
                            );


                        if (field) {

                            handleApplicationColumnSort(
                                field
                            );

                        }

                    }
                );

            }
        );

    }


    /* ============================================================
       APPLICATIONS: ROW CLICK
       ============================================================ */

    function attachApplicationRowClickHandlers() {

        var rows =
            document.querySelectorAll(
                ".nsv-table tbody tr"
            );


        rows.forEach(
            function (row) {

                row.addEventListener(
                    "click",
                    function (event) {

                        if (
                            event.target.classList.contains(
                                "nsv-column-resizer"
                            )
                        ) {

                            return;

                        }


                        var itemId =
                            Number(
                                row.getAttribute(
                                    "data-item-id"
                                )
                            );


                        var item =
                            allApplicationItems.find(
                                function (record) {

                                    return Number(
                                        record.ID
                                    ) ===
                                    itemId;

                                }
                            );


                        if (!item) {

                            return;

                        }


                        selectedApplicationItem =
                            item;


                        openApplicationModal(
                            item
                        );

                    }
                );

            }
        );

    }


    /* ============================================================
       APPLICATIONS: TOGGLE MODE
       ============================================================ */

    function exitApplicationsModeIfActive() {

        if (!applicationsMode) {

            return;

        }


        applicationsMode =
            false;


        var button =
            byId(
                "nsv-applications"
            );


        if (button) {

            button.classList.remove(
                "nsv-update-active"
            );

            button.setAttribute(
                "aria-pressed",
                "false"
            );

        }

    }


    function toggleApplicationsMode() {

        applicationsMode = !applicationsMode;


        var button =
            byId("nsv-applications");

        var dashboardsButton =
            byId("nsv-dashboards");

        var searchInput =
            byId("nsv-search");


        if (applicationsMode) {

            /*
             * Dashboards, Records and Applications are mutually
             * exclusive views. Turning Applications on always
             * turns Dashboards off.
             */
            if (dashboardMode) {

                dashboardMode = false;

                if (dashboardsButton) {

                    dashboardsButton.classList.remove(
                        "nsv-update-active"
                    );

                    dashboardsButton.setAttribute(
                        "aria-pressed",
                        "false"
                    );

                }

            }


            /*
             * Update mode is a Records-table concept (editable
             * rows). It has no meaning on the Applications
             * table, so switch it off if it was left on.
             */
            if (updateMode) {

                updateMode = false;


                var updateButton =
                    byId(
                        "nsv-update-clearance"
                    );

                var updateHint =
                    byId(
                        "nsv-update-hint"
                    );


                if (updateButton) {

                    updateButton.classList.remove(
                        "nsv-update-active"
                    );

                    updateButton.setAttribute(
                        "aria-pressed",
                        "false"
                    );

                }


                if (updateHint) {

                    updateHint.style.display =
                        "none";

                }

            }

            button.classList.add("nsv-update-active");
            button.setAttribute("aria-pressed", "true");

            if (searchInput) {
                searchInput.disabled = false;
            }

            loadApplicationItems().catch(
                function (error) {

                    console.error(
                        "NSV Applications load error:",
                        error
                    );

                    setStatus(
                        "ERROR: " +
                        error.message
                    );

                }
            );

        }
        else {

            button.classList.remove("nsv-update-active");
            button.setAttribute("aria-pressed", "false");

            if (searchInput) {
                searchInput.disabled = false;
            }

            renderTable();

            updatePanelDescription();

            setStatus("Viewing table.");

        }

    }


    /* ============================================================
       APPLICATIONS: MODAL
       ============================================================ */

    function openApplicationModal(
        item
    ) {

        closeModal();


        var fieldsHtml =
            "";


        TABLE_COLUMNS.forEach(
            function (column) {

                fieldsHtml +=
                    buildReadOnlyField(
                        column,
                        formatModalValue(
                            item[
                                column.field
                            ],
                            column.type
                        )
                    );

            }
        );


        var modal =
            document.createElement(
                "div"
            );


        modal.id =
            "nsv-modal-overlay";


        modal.className =
            "nsv-modal-overlay";


        modal.innerHTML = `

            <div
                class="nsv-modal"
                role="dialog"
                aria-modal="true"
                aria-labelledby="nsv-modal-title"
            >

                <div class="nsv-modal-header">

                    <div>

                        <h2
                            id="nsv-modal-title"
                            class="nsv-modal-title"
                        >
                            Application: ${escapeHtml(
                                getFieldValue(
                                    item,
                                    "field_6"
                                )
                            )}
                        </h2>

                        <p class="nsv-modal-description">
                            Review the application, then confirm to move it to the overall NSV register.
                        </p>

                    </div>

                    <button
                        type="button"
                        id="nsv-modal-close"
                        class="nsv-modal-close"
                        aria-label="Close"
                    >
                        &#215;
                    </button>

                </div>

                <div
                    id="nsv-modal-error"
                    class="nsv-modal-error"
                    style="display:none;"
                ></div>

                <div class="nsv-form">

                    <div class="nsv-form-grid">

                        ${fieldsHtml}

                    </div>

                    <div class="nsv-modal-footer">

                        <button
                            type="button"
                            id="nsv-modal-cancel"
                            class="nsv-modal-button nsv-secondary-button"
                        >
                            Close
                        </button>

                        <button
                            type="button"
                            id="nsv-modal-confirm-complete"
                            class="nsv-modal-button nsv-primary-button"
                        >
                            Confirm Complete
                        </button>

                    </div>

                </div>

            </div>

        `;


        document.body.appendChild(
            modal
        );


        document.body.classList.add(
            "nsv-modal-open"
        );


        attachApplicationModalEvents(
            item
        );

    }


    function attachApplicationModalEvents(
        item
    ) {

        var closeButton =
            byId(
                "nsv-modal-close"
            );

        var cancelButton =
            byId(
                "nsv-modal-cancel"
            );

        var confirmButton =
            byId(
                "nsv-modal-confirm-complete"
            );


        if (closeButton) {

            closeButton.addEventListener(
                "click",
                closeModal
            );

        }


        if (cancelButton) {

            cancelButton.addEventListener(
                "click",
                closeModal
            );

        }


        if (confirmButton) {

            confirmButton.addEventListener(
                "click",
                function () {

                    confirmCompleteApplication(
                        item
                    );

                }
            );

        }


        var overlay =
            byId(
                "nsv-modal-overlay"
            );


        if (overlay) {

            overlay.addEventListener(
                "click",
                function (event) {

                    if (
                        event.target ===
                        overlay
                    ) {

                        closeModal();

                    }

                }
            );

        }


        document.addEventListener(
            "keydown",
            modalEscapeHandler
        );

    }


    /* ============================================================
       APPLICATIONS: CONFIRM COMPLETE
       ============================================================ */

    async function confirmCompleteApplication(
        item
    ) {

        if (
            !item ||
            !item.ID
        ) {

            showModalError(
                "The selected application could not be identified."
            );

            return;

        }


        var confirmButton =
            byId(
                "nsv-modal-confirm-complete"
            );

        var cancelButton =
            byId(
                "nsv-modal-cancel"
            );


        if (confirmButton) {

            confirmButton.disabled =
                true;

            confirmButton.classList.add(
                "nsv-button-loading"
            );

            confirmButton.textContent =
                "Moving...";

        }


        if (cancelButton) {

            cancelButton.disabled =
                true;

        }


        /*
         * Copy every field that exists on TABLE_COLUMNS
         * (the schema is shared between the Applications
         * list and the NSV Workbook list) straight across.
         */
        var data = {};

        TABLE_COLUMNS.forEach(
            function (column) {

                var value =
                    item[
                        column.field
                    ];

                data[
                    column.field
                ] =
                    value === undefined
                        ? null
                        : value;

            }
        );


        var createdInWorkbook =
            false;


        try {

            setStatus(
                "Adding record to the NSV register..."
            );


            var createUrl =

                SITE_URL +

                "/_api/web/lists/GetByTitle('" +

                encodeURIComponent(
                    LIST_TITLE
                ) +

                "')/items";


            var createDigest =
                await getRequestDigest();


            await apiPost(
                createUrl,
                JSON.stringify(
                    data
                ),
                {
                    "Content-Type":
                        "application/json;odata=nometadata",

                    "X-RequestDigest":
                        createDigest
                }
            );


            createdInWorkbook =
                true;


            setStatus(
                "Removing application from NSV Applications..."
            );


            var deleteUrl =

                SITE_URL +

                "/_api/web/lists/GetByTitle('" +

                encodeURIComponent(
                    APPLICATIONS_LIST_TITLE
                ) +

                "')/items(" +

                item.ID +

                ")";


            var deleteDigest =
                await getRequestDigest();


            await apiDelete(
                deleteUrl,
                {
                    "X-RequestDigest":
                        deleteDigest
                }
            );


            closeModal();


            selectedApplicationItem =
                null;


            await loadItems();

            await loadApplicationItems();


            setStatus(
                "Successfully moved to overall NSV register."
            );

        }
        catch (error) {

            console.error(
                "Confirm complete error:",
                error
            );


            if (!createdInWorkbook) {

                /*
                 * The record was never created in the
                 * Workbook, so the application is untouched
                 * and it is safe to let the user retry.
                 */
                showModalError(
                    (
                        error &&
                        error.message
                    )
                        ? error.message
                        : "Unable to add the record to the NSV register."
                );


                if (confirmButton) {

                    confirmButton.disabled =
                        false;

                    confirmButton.classList.remove(
                        "nsv-button-loading"
                    );

                    confirmButton.textContent =
                        "Confirm Complete";

                }


                if (cancelButton) {

                    cancelButton.disabled =
                        false;

                }


                setStatus(
                    "ERROR: " +
                    (
                        error &&
                        error.message
                            ? error.message
                            : "Unable to add the record to the NSV register."
                    )
                );

                return;

            }


            /*
             * The record WAS created in the Workbook but the
             * delete from Applications failed. Do not let the
             * button be clicked again — that would create a
             * duplicate Workbook record. Surface a clear,
             * actionable message instead.
             */
            showModalError(
                "The record was added to the NSV register, but could not be removed from " +
                "NSV Applications (ID " +
                item.ID +
                "). Please delete it manually from NSV Applications to avoid a duplicate."
            );


            setStatus(
                "ERROR: record added to NSV register, but the application record (ID " +
                item.ID +
                ") could not be removed automatically. Please delete it manually."
            );


            if (cancelButton) {

                cancelButton.disabled =
                    false;

            }


            await loadItems();

        }

    }


    /* ============================================================
       START
       ============================================================ */

    attachPanelResizeHandler();

    initialise();


})();
