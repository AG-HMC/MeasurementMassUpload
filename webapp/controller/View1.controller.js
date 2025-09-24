sap.ui.define([
  "sap/ui/core/mvc/Controller",
  "sap/ui/export/Spreadsheet",
  "../libs/xlsx",
  "sap/m/MessageToast",
  "sap/m/MessageBox",
  "sap/ui/core/BusyIndicator",
], (Controller, Spreadsheet, xlsx, MessageToast, MessageBox, BusyIndicator) => {
  "use strict";

  return Controller.extend("measuringdocumentmassupload.controller.View1", {
    _ensureSavedColumnSettings: function () {
      try {
        var raw = window.localStorage.getItem("MP_COL_SETTINGS");
        var defaults = this._onGetDefaultColumns() || [];
        // If nothing saved -> write defaults
        if (!raw) {
          window.localStorage.setItem("MP_COL_SETTINGS", JSON.stringify(defaults));
          return defaults;
        }
        // parse saved
        var saved = null;
        try { saved = JSON.parse(raw); } catch (e) { saved = null; }
        if (!Array.isArray(saved)) {
          window.localStorage.setItem("MP_COL_SETTINGS", JSON.stringify(defaults));
          return defaults;
        }

        // Build map of saved by id
        var savedMap = {};
        saved.forEach(function (s) { if (s && s.id) savedMap[s.id] = s; });

        // Merge: for each default, take saved.visible if present (boolean), otherwise default.visible
        var merged = defaults.map(function (def) {
          var s = savedMap[def.id];
          if (!s) return def; // no saved entry -> keep default
          // sanitize visible: only accept true/false; otherwise fallback to default.visible
          var vis = true; // force visible
          var width = (s && s.width && String(s.width).trim()) ? s.width : def.width;
          return { id: def.id, key: def.key, label: def.label, visible: vis, width: width };
        });

        // Also keep any extra saved columns that are not in defaults (preserve user data)
        saved.forEach(function (s) {
          if (s && s.id && !merged.find(function (m) { return m.id === s.id; })) {
            merged.push(s);
          }
        });

        // Save merged back (only if it differs)
        try {
          var mergedRaw = JSON.stringify(merged);
          if (mergedRaw !== raw) window.localStorage.setItem("MP_COL_SETTINGS", mergedRaw);
        } catch (e) { console.warn("Failed to save merged MP_COL_SETTINGS", e); }

        return merged;
      } catch (e) {
        console.error("_ensureSavedColumnSettings failed", e);
        return this._onGetDefaultColumns();
      }
    },

    onInit() {
      this._logsSortDesc = false;
      // ensure batchModel exists
      var comp = this.getOwnerComponent();
      var batchModel = comp.getModel("batchModel");
      if (!batchModel) {
        batchModel = new sap.ui.model.json.JSONModel({ aEmployees: [] });
        comp.setModel(batchModel, "batchModel");
      } else {
        var data = batchModel.getData() || {};
        if (!data.aEmployees) batchModel.setData({ aEmployees: [] });
      }

      // safe initial visibility
      try {
        if (this.byId("tblInput")) this.byId("tblInput").setVisible(true);
        if (this.byId("idTableLayout")) this.byId("idTableLayout").setVisible(false);
      } catch (e) { }
      var CURRENT_VER = "v2";
      var savedRaw = window.localStorage.getItem("MP_COL_SETTINGS_VER");
      if (savedRaw !== CURRENT_VER) {
        // new version: overwrite saved visibility with defaults (or remove saved settings)
        window.localStorage.removeItem("MP_COL_SETTINGS");
        window.localStorage.setItem("MP_COL_SETTINGS_VER", CURRENT_VER);
      }
      try { window.localStorage.removeItem("MP_COL_SETTINGS"); } catch (e) { /* ignore */ }
      // ensure there is always something in localStorage (defaults if missing)
      this._ensureSavedColumnSettings();

      // function to attempt to read and apply saved settings (returns true if applied)
      var applyIfReady = function () {
        try {
          var cfg = this._loadColumnSettings();
          if (!cfg || !Array.isArray(cfg) || cfg.length === 0) return false;

          // check at least one column exists in view
          var foundAny = false;
          for (var i = 0; i < cfg.length; i++) {
            if (!cfg[i] || !cfg[i].id) continue;
            if (this.byId(cfg[i].id)) { foundAny = true; break; }
          }
          if (!foundAny) return false;

          // apply and refresh
          this._applyColumnSettingsToTable(cfg);
          var bm = this.getOwnerComponent().getModel("batchModel");
          if (bm) bm.refresh(true);
          this._colSettingsApplied = true;
          return true;
        } catch (err) {
          console.error("applyIfReady error", err);
          return false;
        }
      }.bind(this);

      // try immediate apply
      if (applyIfReady()) return;

      // attach onAfterRendering delegate to try again (one-time + retries)
      var that = this;
      var tries = 0;
      var maxTries = 8;
      var delegate = {
        onAfterRendering: function () {
          tries++;
          if (that._colSettingsApplied || applyIfReady()) {
            that.getView().removeEventDelegate(delegate);
            return;
          }
          if (tries >= maxTries) {
            that.getView().removeEventDelegate(delegate);
            console.warn("Column settings not applied after retries. Check IDs or localStorage MP_COL_SETTINGS.");
            return;
          }
          // small delay for the next attempt
          setTimeout(function () { if (!that._colSettingsApplied) applyIfReady(); }, 150);
        }
      };
      this.getView().addEventDelegate(delegate, this);
    },
    /* -------------------------
COLUMN SETTINGS / PERSONALIZATION
------------------------- */
    // Sort the ErrorListModel.errItems by State priority
    _sortLogsByState: function (descending) {
      try {
        var oComp = this.getOwnerComponent();
        var oErrM = oComp.getModel("ErrorListModel");
        if (!oErrM) return;

        var data = oErrM.getData() || { errItems: [] };
        var arr = data.errItems || [];

        // priority: FAILED (0) -> SKIPPED (1) -> SUCCESS (2) -> unknown (3)
        var priority = { "FAILED": 0, "SKIPPED": 1, "SUCCESS": 2 };

        arr.sort(function (a, b) {
          var pa = (a && a.State && priority.hasOwnProperty(a.State)) ? priority[a.State] : 3;
          var pb = (b && b.State && priority.hasOwnProperty(b.State)) ? priority[b.State] : 3;
          if (pa === pb) {
            // secondary: show newest first (optional). If you prefer stable original order, return 0.
            try {
              var ta = a && a.Timestamp ? new Date(a.Timestamp).getTime() : 0;
              var tb = b && b.Timestamp ? new Date(b.Timestamp).getTime() : 0;
              return tb - ta;
            } catch (e) { return 0; }
          }
          return descending ? (pb - pa) : (pa - pb);
        });

        data.errItems = arr;
        oErrM.setData(data);
        oErrM.refresh(true);
      } catch (e) {
        console.error("_sortLogsByState failed", e);
      }
    },

    // toggled by the header sort button - public method for view binding
    onPressToggleLogSort: function () {
      // store state on controller instance
      this._logsSortDesc = !!this._logsSortDesc;        // ensure boolean
      this._logsSortDesc = !this._logsSortDesc;        // flip each press
      // call sorter: note default desired order is FAILED -> SKIPPED -> SUCCESS (descending=false)
      this._sortLogsByState(this._logsSortDesc);

      // optional: update button icon/tooltip to indicate direction
      try {
        var btn = this.byId("btnSortStatus");
        if (btn) {
          btn.setTooltip(this._logsSortDesc ? "Sort: lower priority first" : "Sort: FAILED first");
          // change icon if you want:
          btn.setIcon(this._logsSortDesc ? "sap-icon://sort-descending" : "sap-icon://sort-ascending");
        }
      } catch (e) { /* ignore UI update errors */ }
    },

    _onGetDefaultColumns: function () {
      return [
        { id: "colSelect", key: "select", label: "Select", visible: true, width: "2rem" },
        { id: "colMeasuringPoint", key: "MeasuringPoint", label: "Measuring Point", visible: true, width: "8rem" },
        { id: "colDescription", key: "MeasuringPointDescription", label: "Description", visible: true, width: "10rem" },
        { id: "colPosition", key: "MeasuringPointPositionNumber", label: "Position Number", visible: true, width: "10rem" },
        { id: "colCounter", key: "Counter", label: "Counter/Reading", visible: true, width: "5rem" },
        { id: "colDifference", key: "Difference", label: "Difference", visible: true, width: "5rem" },
        { id: "colPostingDate", key: "PostingDate", label: "Posting Date", visible: true, width: "8rem" },
        { id: "colText", key: "MeasurementDocumentText", label: "Text", visible: true, width: "20rem" },
        { id: "colReadyBy", key: "ReadyBy", label: "Read By", visible: true, width: "8rem" }
      ];
    },

    _applyColumnSettingsToTable: function (aColCfg) {
      // aColCfg = array of {id, visible, width}
      aColCfg.forEach(function (c) {
        try {
          var oCol = this.byId(c.id);
          if (oCol) {
            oCol.setVisible(Boolean(c.visible));
            // set width on Column control
            if (c.width) {
              oCol.setWidth(String(c.width));
            }
          }
        } catch (e) { /* ignore per-column errors */ }
      }.bind(this));

      // if items or model need refresh
      var batchModel = this.getOwnerComponent().getModel("batchModel");
      if (batchModel) batchModel.refresh(true);
    },
    _loadColumnSettings: function () {
      try {
        var raw = window.localStorage.getItem("MP_COL_SETTINGS");
        if (!raw) return null;
        var cfg = JSON.parse(raw);
        return cfg;
      } catch (e) {
        return null;
      }
    },
    _saveColumnSettings: function (cfg) {
      try {
        window.localStorage.setItem("MP_COL_SETTINGS", JSON.stringify(cfg));
      } catch (e) {
        console.error("Failed to save column settings", e);
      }
    }, _onBuildSettingsDialogContent: function (aColCfg) {
      // returns a VBox (sap.m.VBox) containing controls for each column
      var oVBox = new sap.m.VBox({ renderType: "Bare", width: "100%" });
      aColCfg.forEach(function (col) {
        var oHBox = new sap.m.HBox({ justifyContent: "SpaceBetween", alignItems: "Center", width: "100%" });
        var oChk = new sap.m.CheckBox({
          selected: !!col.visible,
          text: col.label,
          // let it expand instead of fixed 60%
          width: "auto",
          wrapping: true, // 👈 ensures long text wraps
          customData: [new sap.ui.core.CustomData({ key: "colId", value: col.id })]
        });

        var oInput = new sap.m.Input({
          value: col.width || "",
          placeholder: "e.g. 120px or 8rem",
          width: "35%"
        });
        oInput.addCustomData(new sap.ui.core.CustomData({ key: "colId", value: col.id }));

        oHBox.addItem(oChk);
        oHBox.addItem(oInput);
        oVBox.addItem(oHBox);
      });
      return oVBox;
    },

    onPressColumnSettings: function () {
      // build/restore dialog
      var that = this;
      var saved = this._loadColumnSettings();
      var cols = saved || this._onGetDefaultColumns();

      // create dialog if not already created
      if (!this._oColSettingsDialog) {
        // content container that will scroll if there are many columns
        var oScroll = new sap.m.ScrollContainer({
          height: "50vh",       // 50% of viewport height (use px if you prefer, e.g. "400px")
          width: "100%",
          vertical: true,
          horizontal: false,
          focusable: true
        });

        this._oColSettingsDialog = new sap.m.Dialog({
          title: "Column settings",
          contentWidth: "600px",
          content: [oScroll],
          beginButton: new sap.m.Button({
            text: "Apply & Save",
            press: function () {
              try {
                // find ScrollContainer (first content entry of dialog)
                var oDialog = that._oColSettingsDialog;
                var oScroll = oDialog && oDialog.getContent && oDialog.getContent()[0];
                // the VBox with HBoxes is inside the ScrollContainer's content
                var vbox = (oScroll && oScroll.getContent && oScroll.getContent()[0]) || null;

                var newCfg = [];
                if (vbox && vbox.getItems) {
                  var items = vbox.getItems(); // array of HBox controls
                  items.forEach(function (hbox) {
                    var children = hbox.getItems ? hbox.getItems() : [];
                    var chk = children[0]; // expected CheckBox
                    var inp = children[1]; // expected Input
                    // attempt to get colId from CustomData (works for CheckBox or Input)
                    var colId = null;
                    try {
                      var cd = (chk && chk.getCustomData && chk.getCustomData()[0]) || (inp && inp.getCustomData && inp.getCustomData()[0]);
                      colId = cd && cd.getValue ? cd.getValue() : null;
                    } catch (e) { colId = null; }

                    newCfg.push({
                      id: colId || null,
                      label: (chk && chk.getText && chk.getText()) || "",
                      visible: !!(chk && chk.getSelected && chk.getSelected()),
                      width: (inp && inp.getValue && inp.getValue()) || ""
                    });
                  });
                }

                // apply only those with an id found
                var cfgToApply = newCfg.filter(function (c) { return !!c.id; });
                that._applyColumnSettingsToTable(cfgToApply);
                that._saveColumnSettings(cfgToApply);
                oDialog.close();
              } catch (err) {
                console.error("Apply & Save error:", err);
                sap.m.MessageBox.error("Failed to apply column settings: " + (err && err.message ? err.message : err));
              }
            }
          }),
          endButton: new sap.m.Button({
            text: "Cancel",
            press: function () { that._oColSettingsDialog.close(); }
          }),
          afterClose: function () { /* keep dialog for reuse */ }
        });
        this.getView().addDependent(this._oColSettingsDialog);
      }

      // set dialog content (rebuild to reflect current config)

      // Build the content (VBox) fresh each time from current config
      try {
        // Get current config (saved or defaults)
        var saved = this._loadColumnSettings();
        var cols = saved || this._onGetDefaultColumns();

        // build VBox content using your helper (re-use if you have one)
        var vboxContent = this._onBuildSettingsDialogContent(cols);
        // safe access to ScrollContainer (it was created above)
        var oScroll = this._oColSettingsDialog.getContent()[0];

        if (oScroll && oScroll.removeAllContent && oScroll.addContent) {
          // replace scroll content with the newly built vbox
          oScroll.removeAllContent();
          oScroll.addContent(vboxContent);
        } else {
          // fallback: put vbox directly into dialog content (should not happen)
          this._oColSettingsDialog.removeAllContent();
          this._oColSettingsDialog.addContent(vboxContent);
        }

        // open dialog
        this._oColSettingsDialog.open();
      } catch (err) {
        console.error("Opening settings dialog failed:", err);
        sap.m.MessageBox.error("Failed to open column settings: " + (err && err.message ? err.message : err));
      }
    },
    // fetch MP details using OData V2 model
    _fetchMeasuringPointDetails: function (mpId) {
      var that = this;
      return new Promise(function (resolve) {
        var oModel = that.getOwnerComponent().getModel("MPModel");
        if (!oModel) {
          resolve({ found: false, error: "MPModel not available" });
          return;
        }
        var sKey = encodeURIComponent(String(mpId));
        var sPath = "/zc_measuringpointdata('" + sKey + "')";

        oModel.read(sPath, {
          success: function (data) {
            resolve({ found: true, data: data });
          },
          error: function (err) {
            var msg = (err && err.message) ? err.message : "Error reading MP";
            resolve({ found: false, error: msg });
          }
        });
      });
    },
    _getLatestMeasurementForMP: function (measuringPoint) {
      var that = this;
      return new Promise(function (resolve) {
        try {
          console.info("_getLatestMeasurementForMP start for MP:", measuringPoint);
          if (!measuringPoint) {
            console.warn("_getLatestMeasurementForMP: no measuringPoint provided");
            return resolve(null);
          }

          var oComp = that.getOwnerComponent();
          var globalModel = oComp.getModel("") || null;
          var sBearer = "";
          try { if (globalModel) sBearer = globalModel.getProperty("/auth/token") || globalModel.getProperty("/bearerToken") || ""; } catch (e) { sBearer = ""; }

          // Prefer using the OData model (safer in ABAP environment)
          var oDataModel = oComp.getModel(); // default model (maybe OData)
          if (oDataModel && typeof oDataModel.read === "function") {
            try {
              // Build path/query - use system's OData V4 read if supported by your model
              var filter = "MeasuringPoint eq '" + String(measuringPoint) + "'";
              var path = "/MeasurementDocument"; // try relative entity; fallback to full URI if needed
              // If your default model is not the correct service, skip to fetch
              console.info("_getLatestMeasurementForMP: trying oDataModel.read with filter:", filter);
              oDataModel.read(path, {
                urlParameters: {
                  "$filter": filter,
                  "$orderby": "MsmtRdngDate desc,MsmtRdngTime desc",
                  "$top": "1",
                  "$select": "MeasurementCounterReading,MeasurementReading,MsmtRdngDate,MsmtRdngTime"
                },
                success: function (data) {
                  try {
                    console.info("_getLatestMeasurementForMP: oDataModel.read success", data);
                    var arr = data && data.value ? data.value : (Array.isArray(data) ? data : null);
                    if (arr && arr.length) {
                      var rec = arr[0];
                      var last = null;
                      if (rec.MeasurementCounterReading !== undefined && rec.MeasurementCounterReading !== null) last = Number(rec.MeasurementCounterReading);
                      else if (rec.MeasurementReading !== undefined && rec.MeasurementReading !== null) last = Number(rec.MeasurementReading);
                      if (!isNaN(last)) { console.info("_getLatestMeasurementForMP: last reading from oDataModel:", last); return resolve(last); }
                    }
                  } catch (e) { console.warn("_getLatestMeasurementForMP: parsing oDataModel response failed", e); }
                  return resolve(null);
                },
                error: function (err) {
                  console.warn("_getLatestMeasurementForMP: oDataModel.read error - will fallback to fetch", err);
                  // continue to fetch fallback below
                  fetchFallback();
                }
              });
              return;
            } catch (e) {
              console.warn("_getLatestMeasurementForMP: oDataModel.read attempt threw, fallback to fetch", e);
              // proceed to fetch fallback
            }
          }

          // fallback: use fetch to OData V4 endpoint (full absolute path as used in your POST)
          function fetchFallback() {
            try {
              var mpEncoded = encodeURIComponent(String(measuringPoint));
              var base = "/sap/opu/odata4/sap/api_measurementdocument/srvd_a2x/sap/MeasurementDocument/0001/MeasurementDocument";
              var q = base + "?$filter=MeasuringPoint eq '" + mpEncoded + "'&$orderby=MsmtRdngDate desc,MsmtRdngTime desc&$top=1&$select=MeasurementCounterReading,MeasurementReading,MsmtRdngDate,MsmtRdngTime";
              console.info("_getLatestMeasurementForMP: fetch URL:", q);

              var headers = { "Accept": "application/json" };
              if (sBearer) headers["Authorization"] = "Bearer " + sBearer;

              fetch(q, { method: "GET", headers: headers }).then(function (resp) {
                console.info("_getLatestMeasurementForMP: fetch status", resp.status);
                if (!resp.ok) {
                  console.warn("_getLatestMeasurementForMP: fetch not ok, resolving null");
                  return resolve(null);
                }
                return resp.text().then(function (txt) {
                  var body = null;
                  try { body = txt && txt.length ? JSON.parse(txt) : null; } catch (e) { body = txt; }
                  console.info("_getLatestMeasurementForMP: fetch body", body);
                  var arr = (body && body.value && Array.isArray(body.value)) ? body.value : null;
                  if (arr && arr.length) {
                    var rec = arr[0];
                    var last = null;
                    if (rec.MeasurementCounterReading !== undefined && rec.MeasurementCounterReading !== null) last = Number(rec.MeasurementCounterReading);
                    else if (rec.MeasurementReading !== undefined && rec.MeasurementReading !== null) last = Number(rec.MeasurementReading);
                    if (!isNaN(last)) { console.info("_getLatestMeasurementForMP: lastReading fetched:", last); return resolve(last); }
                  }
                  return resolve(null);
                });
              }).catch(function (err) {
                console.error("_getLatestMeasurementForMP: fetch error", err);
                return resolve(null);
              });
            } catch (ex) {
              console.error("_getLatestMeasurementForMP: exception", ex);
              return resolve(null);
            }
          }

          // call fallback if we reach here
          fetchFallback();

        } catch (errOuter) {
          console.error("_getLatestMeasurementForMP outer exception:", errOuter);
          return resolve(null);
        }
      });
    },

    // enrich Excel rows with description and position
    _enrichRowsWithMP: async function (aRows) {
      var batchModel = this.getOwnerComponent().getModel("batchModel");

      for (var i = 0; i < aRows.length; i++) {
        var row = aRows[i];
        var mp = row.MeasuringPoint || row["Measuring Point"];

        row.MeasuringPointDescription = "Loading...";
        row.MeasuringPointPositionNumber = "";
        if (batchModel) batchModel.refresh(true);

        try {
          var res = await this._fetchMeasuringPointDetails(mp);
          if (res.found && res.data) {
            row.MeasuringPointDescription = res.data.MeasuringPointDescription || "";
            row.MeasuringPointPositionNumber = res.data.MeasuringPointPositionNumber || "";

            // --- ADD: UoM from CDS service (field name observed: MeasurementRangeUnit) ---
            // check the exact property name returned by your service (in your screenshot it's MeasurementRangeUnit)
            row.MeasuringPointUoM = res.data.MeasurementRangeUnit || res.data.MeasuringPointUoM || "H";
          } else {
            row.MeasuringPointDescription = "Not found";
            row.MeasuringPointPositionNumber = "";
            row.MeasuringPointUoM = "H"; // fallback
          }
        } catch (e) {
          row.MeasuringPointDescription = "Error";
          row.MeasuringPointPositionNumber = "";
        } finally {
          if (batchModel) batchModel.refresh(true);
        }
      }
    },

    _logResult: function (payload, message, state) {
      try {
        var oComp = this.getOwnerComponent();
        var errModel = oComp.getModel("ErrorListModel");
        if (!errModel) {
          errModel = new sap.ui.model.json.JSONModel({ errItems: [] });
          oComp.setModel(errModel, "ErrorListModel");
        }
        var data = errModel.getData() || { errItems: [] };
        data.errItems = data.errItems || [];
        // Prefer MeasurementReading (new payload) then MeasurementCounterReading then MeasurementReading (older usage)
        // var val = "";
        // if (payload && payload.MeasurementReading !== undefined && payload.MeasurementReading !== null && payload.MeasurementReading !== "") {
        //   val = payload.MeasurementReading;
        // } else if (payload && payload.MeasurementCounterReading !== undefined && payload.MeasurementCounterReading !== null) {
        //   val = payload.MeasurementCounterReading;
        // } else if (payload && payload.Counter !== undefined && payload.Counter !== null) {
        //   val = payload.Counter;
        // }
        // --- FIX: pick value correctly ---
        var val = "";
        if (payload) {
          if (payload.MeasurementReading !== undefined && payload.MeasurementReading !== null && payload.MeasurementReading !== "") {
            val = payload.MeasurementReading;
          } else if (payload.MsmtCounterReadingDifference !== undefined && payload.MsmtCounterReadingDifference !== null && payload.MsmtCounterReadingDifference !== "") {
            val = payload.MsmtCounterReadingDifference;   // 👈 show difference
          } else if (payload.MeasurementCounterReading !== undefined && payload.MeasurementCounterReading !== null) {
            val = payload.MeasurementCounterReading;
          } else if (payload.Counter !== undefined && payload.Counter !== null) {
            val = payload.Counter;
          }
        }

        data.errItems.push({
          Equipment: payload.MeasuringPoint || "",
          Value: val,
          // RawValue: JSON.stringify(payload),
          ErrorText: message || "",
          State: state || "FAILED",
          Timestamp: new Date().toISOString()
        });
        errModel.setData(data);
        // ensure logs area visible
        try { this.byId("idTableLayout") && this.byId("idTableLayout").setVisible(true); } catch (e) { }
      } catch (e) {
        console.error("_logResult error:", e);
      }
    },

    /* <<< ADD: helper to extract friendly SAP error message (paste right after _logResult) >>> */
    _extractSapErrorMessage: function (body, txt) {
      try {
        if (body && body.SAP__Messages && Array.isArray(body.SAP__Messages) && body.SAP__Messages.length) {
          var m = body.SAP__Messages[0];
          return m.Message || m.MessageText || (typeof m === "string" ? m : JSON.stringify(m));
        }
        if (body && body.error && body.error.message) {
          if (typeof body.error.message === "string") return body.error.message;
          if (body.error.message.value) return body.error.message.value;
          if (body.error.message.Message) return body.error.message.Message;
        }
        if (body && body.error && body.error.innererror) {
          var ie = body.error.innererror;
          if (ie.errordetails && Array.isArray(ie.errordetails) && ie.errordetails.length) {
            var d = ie.errordetails[0];
            return d.message || d.Message || d.messageValue || JSON.stringify(d);
          }
          if (ie.ErrorDetails && Array.isArray(ie.ErrorDetails) && ie.ErrorDetails.length) {
            var dd = ie.ErrorDetails[0];
            return dd.message || dd.Message || JSON.stringify(dd);
          }
        }
        if (txt && typeof txt === "string") {
          var mreg = /"message"\s*:\s*"([^"]+)"/i;
          var m = mreg.exec(txt);
          if (m && m[1]) return m[1];
          var mtreg = /MessageText['"]?\s*[:=]\s*["']([^"']+)["']/i;
          var mm = mtreg.exec(txt);
          if (mm && mm[1]) return mm[1];
          var trimmed = txt.trim();
          if (trimmed.length < 400) return trimmed;
          var first = trimmed.split(/[\r\n\.]{1,2}/)[0];
          if (first && first.length < 400) return first + "...";
        }
      } catch (e) {
        console.debug("_extractSapErrorMessage parsing failed", e);
      }
      return (txt && typeof txt === "string") ? (txt.substr(0, 600) + (txt.length > 600 ? "..." : "")) : "Unknown error from server";
    },

    _postServicePromise: function (payload) {
      var that = this;
      return new Promise(function (resolve, reject) {
        try {
          var oComp = that.getOwnerComponent();
          var globalModel = oComp.getModel("");
          var sBearer = "";

          try {
            if (globalModel) {
              // adjust these keys to where you store the token if different
              sBearer = globalModel.getProperty("/auth/token") || globalModel.getProperty("/bearerToken") || "";
            }
          } catch (e) { sBearer = ""; }

          // MeasurementDocument endpoint (OData V4 create)
          var sUrl = "/sap/opu/odata4/sap/api_measurementdocument/srvd_a2x/sap/MeasurementDocument/0001/MeasurementDocument";

          // Error list model (ensure exists)
          var errModel = oComp.getModel("ErrorListModel");
          if (!errModel) {
            errModel = new sap.ui.model.json.JSONModel({ errItems: [] });
            oComp.setModel(errModel, "ErrorListModel");
          }

          // Build headers for CSRF GET
          var headersGet = { "X-CSRF-Token": "Fetch", "Accept": "application/json" };
          if (sBearer) headersGet["Authorization"] = "Bearer " + sBearer;

          // 1) GET CSRF (works even if CSRF not strictly required when using Bearer)
          fetch(sUrl, {
            method: "GET",
            headers: headersGet,
            //  credentials: "include"
          }).then(function (resp) {
            if (!resp.ok) {
              return resp.text().then(function (txt) {
                var msg = "CSRF GET failed HTTP " + resp.status + " " + (txt || "");
                throw new Error(msg);
              });
            }
            // try to read token (may be null)
            var token = resp.headers.get("x-csrf-token") || resp.headers.get("X-CSRF-Token") || null;
            return token;
          }).then(function (csrfToken) {
            // Prepare POST headers
            var headersPost = { "Content-Type": "application/json", "Accept": "application/json" };
            //   if (sBearer) headersPost["Authorization"] = "Bearer " + sBearer;
            if (csrfToken) headersPost["X-CSRF-Token"] = csrfToken;

            // 2) POST payload
            return fetch(sUrl, {
              method: "POST",
              headers: headersPost,
              //credentials: "include",
              body: JSON.stringify(payload)
            }).then(function (postResp) {
              return postResp.text().then(function (txt) {
                var body = null;
                try { body = txt && txt.length ? JSON.parse(txt) : null; } catch (e) { body = txt; }

                if (postResp.ok) {
                  // success - log and resolve
                  var successMsg = "SUCCESS" + (body && body.MeasurementDocument ? (" - Doc: " + body.MeasurementDocument) : "");
                  that._logResult(payload, successMsg, "SUCCESS");
                  resolve(body || { status: postResp.status });
                } else {
                  var shortMsg = that._extractSapErrorMessage(body, txt); // <<< CHANGE: use helper
                  var errMsg = "HTTP " + postResp.status + " - " + shortMsg;
                  that._logResult(payload, errMsg, "FAILED");
                  reject(new Error(errMsg));
                }
              });
            });
          }).catch(function (err) {
            var message = (err && err.message) ? err.message : String(err);
            var extracted = message;
            try {
              var maybeBody = null;
              try { maybeBody = JSON.parse(message); } catch (e) { maybeBody = null; }
              if (maybeBody) {
                extracted = that._extractSapErrorMessage(maybeBody, message); // <<< CHANGE: try to extract human msg
              } else {
                extracted = message;
              }
            } catch (e2) {
              extracted = message;
            }
            that._logResult(payload, "ERROR: " + extracted, "FAILED");
            reject(err);
          });

        } catch (ex) {
          that._logResult(payload, "Client exception: " + (ex && ex.message ? ex.message : String(ex)), "FAILED");
          reject(ex);
        }
      });
    },

    onPressSubmitBatch: async function () {

      var that = this;
      try {
        var oTable = this.byId("tblInput");
        if (!oTable) {
          sap.m.MessageBox.error("Table (tblInput) not found");
          return;
        }

        // Read selected items (only those will be posted)
        var aSelected = oTable.getSelectedItems();
        if (!aSelected || aSelected.length === 0) {
          sap.m.MessageToast.show("Please select one or more rows to post");
          return;
        }

        // Map selected binding contexts to row objects
        var aRows = [];
        aSelected.forEach(function (item) {
          var ctx = item.getBindingContext("batchModel");
          if (ctx) aRows.push(ctx.getObject());
        });

        if (!aRows.length) {
          sap.m.MessageToast.show("No valid selected rows found");
          return;
        }
        // ======= INSERT: before starting uploads: show logs, hide input table =======
        try {
          this.byId("idTableLayout") && this.byId("idTableLayout").setVisible(true);
          this.byId("tblInput") && this.byId("tblInput").setVisible(false);
        } catch (e) { /* ignore */ }
        // ===========================================================================
        // Helpers to format date/time
        /*var toYYYYMMDD = function (s) {   
          if (!s) {
            var d0 = new Date();
            return d0.getFullYear() + "-" + String(d0.getMonth() + 1).padStart(2, "0") + "-" + String(d0.getDate()).padStart(2, "0");
          }
          s = String(s).trim();
          if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
          if (/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(s)) {
            var p = s.split(/[-/]/); return p[2] + "-" + p[1].padStart(2, "0") + "-" + p[0].padStart(2, "0");
          }
          var d = new Date(s);
          if (!isNaN(d.getTime())) return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
          var d2 = new Date();
          return d2.getFullYear() + "-" + String(d2.getMonth() + 1).padStart(2, "0") + "-" + String(d2.getDate()).padStart(2, "0");
        };*/

        var toYYYYMMDD = function (s) {
          // Default: today
          var today = new Date();
          var defaultStr = today.getFullYear() + "-" +
            String(today.getMonth() + 1).padStart(2, "0") + "-" +
            String(today.getDate()).padStart(2, "0");

          if (!s) return defaultStr;
          s = String(s).trim();

          // Already in YYYY-MM-DD
          if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
            return s;
          }

          // SAP style: /Date(....)/
          if (s.indexOf("/Date(") === 0) {
            var ms = parseInt(s.replace("/Date(", "").replace(")/", ""), 10);
            var d0 = new Date(ms);
            return d0.getFullYear() + "-" +
              String(d0.getMonth() + 1).padStart(2, "0") + "-" +
              String(d0.getDate()).padStart(2, "0");
          }

          // dd-mm-yyyy or dd/mm/yyyy
          if (/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(s)) {
            var p1 = s.split(/[-/]/);
            return p1[2] + "-" + p1[1].padStart(2, "0") + "-" + p1[0].padStart(2, "0");
          }

          // dd.mm.yyyy
          if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) {
            var p2 = s.split(".");
            return p2[2] + "-" + p2[1].padStart(2, "0") + "-" + p2[0].padStart(2, "0");
          }

          // ddmmyyyy (no delimiters, distinguish from yyyymmdd)
          if (/^\d{8}$/.test(s)) {
            var day = s.substring(0, 2);
            var month = s.substring(2, 4);
            var year = s.substring(4, 8);
            if (parseInt(day, 10) <= 31 && parseInt(month, 10) <= 12) {
              return year + "-" + month + "-" + day;
            }
            // Otherwise assume already YYYYMMDD
            return s.substring(0, 4) + "-" + s.substring(4, 6) + "-" + s.substring(6, 8);
          }

          // Excel serial number (e.g., 45921 → 2025-09-22)
          if (/^\d+$/.test(s)) {
            var n = Number(s);
            if (n > 59) n -= 1; // Excel leap year bug (1900 treated as leap)
            var epoch = new Date(Date.UTC(1899, 11, 31));
            epoch.setUTCDate(epoch.getDate() + n);
            return epoch.getFullYear() + "-" +
              String(epoch.getMonth() + 1).padStart(2, "0") + "-" +
              String(epoch.getDate()).padStart(2, "0");
          }

          // Fallback generic parse
          var d3 = new Date(s);
          if (!isNaN(d3.getTime())) {
            return d3.getFullYear() + "-" +
              String(d3.getMonth() + 1).padStart(2, "0") + "-" +
              String(d3.getDate()).padStart(2, "0");
          }

          // Final fallback: today
          return defaultStr;
        };

        var systemTimeHHMMSS = function () {
          var d = new Date();
          return String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0") + ":" + String(d.getSeconds()).padStart(2, "0");
        };
        var buildPayloadFromRow = function (row) {
          var getVal = function (names, def) {
            def = (def === undefined) ? "" : def;
            for (var i = 0; i < names.length; i++) {
              var k = names[i];
              if (Object.prototype.hasOwnProperty.call(row, k) && row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== "") {
                return row[k];
              }
            }
            return def;
          };

          var parseNum = function (v) {
            if (v === "" || v === null || v === undefined) return null;
            if (typeof v === "number") return v;
            var n = Number(v);
            if (!isNaN(n)) return n;
            // clean strings, allow minus, digits, dot, comma
            var cleaned = String(v).replace(/[^0-9\.\-\,]/g, "").replace(",", ".");
            n = Number(cleaned);
            return isNaN(n) ? null : n;
          };

          // read common fields (accept alternate header names)
          var mp = String(getVal(["Measuring Point", "MeasuringPoint", "Measuring_Point", "Equipment"], "")).trim();
          var readingRaw = getVal(["Reading", "MeasurementReading", "MeasurementCounterReading", "Counter", "Value"], "");
          var diffRaw = getVal(["Difference", "MsmtCounterReadingDifference", "CounterDifference"], "");
          var postDate = toYYYYMMDD(getVal(["Posting Date", "MsmtRdngDate", "Date", "PostingDate"], ""));
          var text = String(getVal(["MeasurementDocumentText", "Text", "LongText"], "Reading Taken")).trim();
          var user = String(getVal(["Read By", "ReadBy", "Ready By", "ReadyBy", "MsmtRdngByUser", "User"], "USER")).trim();

          var readingNum = parseNum(readingRaw);
          var diffNum = parseNum(diffRaw);

          // base/common payload fields
          var payload = {
            MeasuringPoint: mp,
            MsmtRdngDate: postDate,
            MsmtRdngTime: systemTimeHHMMSS(),
            MsmtRdngStatus: String(getVal(["MsmtRdngStatus", "Status"], "1")),
            MeasurementDocumentText: text,
            MsmtRdngByUser: user,
            MsmtIsDoneAfterTaskCompltn: !!getVal(["MsmtIsDoneAfterTaskCompltn", "MsmtIsDoneAfterTaskCompletion", "IsDoneAfterTask"], false),
            // optional UoM if you enriched it earlier:
            MeasurementReadingEntryUoM: (row && row.MeasuringPointUoM) ? row.MeasuringPointUoM : undefined
          };

          // DECIDE: difference OR reading (difference takes precedence if present)
          if (diffNum !== null) {
            // DIFFERENCE MODE (send diff flag + diff value). Do NOT send MeasurementReading.
            payload.MsmtCntrReadingDiffIsEntered = true;                 // required flag per API
            payload.MsmtCounterReadingDifference = diffNum;              // numeric diff value
            // remove undefined props to keep payload clean
            if (payload.MeasurementReading !== undefined) delete payload.MeasurementReading;
          } else if (readingNum !== null) {
            // READING MODE (send reading). Do NOT send MsmtCounterReadingDifference.
            payload.MeasurementReading = readingNum;                     // or MeasurementCounterReading if your API expects that
            payload.MsmtCntrReadingDiffIsEntered = false;               // optional; API usually accepts absence as well
            if (payload.MsmtCounterReadingDifference !== undefined) delete payload.MsmtCounterReadingDifference;
          } else {
            // Neither reading nor difference present — keep payload minimal. Server will validate & return an error.
            payload.MsmtCntrReadingDiffIsEntered = false;
          }

          // Remove any undefined fields (so we do not send keys with undefined)
          Object.keys(payload).forEach(function (k) {
            if (payload[k] === undefined) delete payload[k];
          });

          // DEBUG: inspect payload in console (remove/comment out in production)
          try { console.info("Payload for MP:", mp, payload); } catch (e) { }

          return payload;

        };

        // Show global BusyIndicator
        sap.ui.core.BusyIndicator.show(0);

        // ---------- REPLACE existing posting loop with this (no popup, just log SKIPPED) ----------
        for (var i = 0; i < aRows.length; i++) {
          var row = aRows[i];

          // UI feedback
          row._uploadStatus = "Validating...";
          var batchModel = this.getOwnerComponent().getModel("batchModel");
          if (batchModel) batchModel.refresh(true);

          // build payload
          var payload = buildPayloadFromRow(row);
          var readingNum = Number(payload.MeasurementCounterReading || payload.MeasurementReading || 0) || 0;
          var mp = payload.MeasuringPoint || "";

          try {
            // otherwise proceed to post
            row._uploadStatus = "Uploading...";
            if (batchModel) batchModel.refresh(true);

            await this._postServicePromise(payload);
            row._uploadStatus = "SUCCESS";
          } catch (err) {
            row._uploadStatus = "FAILED";
            console.error("Upload error for MeasuringPoint", payload.MeasuringPoint, err);
            // _postServicePromise already logs details; but ensure UI refreshed
          } finally {
            if (batchModel) batchModel.refresh(true);
            // small throttle
            await new Promise(function (res) { setTimeout(res, 150); });
          }
        }
        // ---------- end loop ----------

        sap.ui.core.BusyIndicator.hide();
        // ======= INSERT: after finishing uploads: keep logs visible and hide input table =======
        try {
          this.byId("idTableLayout") && this.byId("idTableLayout").setVisible(true);
          this.byId("tblInput") && this.byId("tblInput").setVisible(false);
        } catch (e) { /* ignore */ }
        // ===========================================================================
        sap.m.MessageToast.show("Selected records processed. Check Upload Logs.");
        // refresh logs model
        var errM = this.getOwnerComponent().getModel("ErrorListModel");
        if (errM) errM.refresh(true);
        this._sortLogsByState(false);

      } catch (ex) {
        sap.ui.core.BusyIndicator.hide();
        console.error("onPressSubmitBatch error:", ex);
        sap.m.MessageBox.error("Unexpected error: " + (ex && ex.message ? ex.message : String(ex)));
      }
    },

    _toYYYYMMDD: function (s) {
      if (!s) return null;
      s = String(s).trim();

      // Already YYYYMMDD
      if (/^\d{8}$/.test(s)) {
        return s;
      }

      // SAP style: /Date(....)/
      if (s.indexOf("/Date(") === 0) {
        var ms = parseInt(s.replace("/Date(", "").replace(")/", ""), 10);
        var d = new Date(ms);
        return d.getFullYear().toString() +
          String(d.getMonth() + 1).padStart(2, "0") +
          String(d.getDate()).padStart(2, "0");
      }

      // dd-mm-yyyy or dd/mm/yyyy
      if (/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(s)) {
        var parts = s.split(/[-/]/);
        return parts[2] + parts[1].padStart(2, "0") + parts[0].padStart(2, "0");
      }

      // dd.mm.yyyy
      if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) {
        var parts2 = s.split(".");
        return parts2[2] + parts2[1].padStart(2, "0") + parts2[0].padStart(2, "0");
      }

      // ddmmyyyy (no delimiters, distinguish from yyyymmdd)
      if (/^\d{8}$/.test(s)) {
        var day = s.substring(0, 2);
        var month = s.substring(2, 4);
        var year = s.substring(4, 8);
        if (parseInt(day, 10) <= 31 && parseInt(month, 10) <= 12) {
          return year + month + day;
        }
        // Otherwise assume already YYYYMMDD
        return s;
      }

      // Excel serial number (e.g., 45921 → 20250922)
      if (/^\d+$/.test(s)) {
        var n = Number(s);
        if (n > 59) n -= 1; // Excel leap year bug (1900 treated as leap)
        var epoch = new Date(Date.UTC(1899, 11, 31)); // Excel base date
        epoch.setUTCDate(epoch.getUTCDate() + n);
        return epoch.getUTCFullYear().toString() +
          String(epoch.getUTCMonth() + 1).padStart(2, "0") +
          String(epoch.getUTCDate()).padStart(2, "0");
      }

      // Fallback generic parse
      var d3 = new Date(s);
      if (!isNaN(d3.getTime())) {
        return d3.getUTCFullYear().toString() +
          String(d3.getUTCMonth() + 1).padStart(2, "0") +
          String(d3.getUTCDate()).padStart(2, "0");
      }

      return null; // not parseable
    }
    , // ✅ closing brace was missing
    onSelectFile: function (oEvent) {
      this._import(oEvent.getParameter("files") && oEvent.getParameter("files")[0]);
    },

    _import: async function (file) {
      var that = this;
      if (!(file && window.FileReader)) {
        sap.m.MessageBox.error("File API not supported in this browser");
        return;
      }
      sap.ui.core.BusyIndicator.show(0);   // <<< show busy
      var reader = new FileReader();

      reader.onload = async function (e) {
        try {
          var data = e.target.result;
          var workbook = XLSX.read(data, { type: "binary" });

          // use first sheet by default (change index if needed)
          var sheetName = workbook.SheetNames[0];
          var aResults = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName] || {});

          // aResults = aResults.map(function (row) {

          //   // keep original values as fallback
          //   var normalized = {};
          //   Object.keys(row || {}).forEach(function (k) { normalized[k] = row[k]; });

          //   var lookup = function(names) {
          //     for (var i = 0; i < names.length; i++) {
          //       var n = names[i];
          //       if (Object.prototype.hasOwnProperty.call(row, n) && row[n] !== undefined && row[n] !== null && String(row[n]).trim() !== "") {
          //         return row[n];
          //       }
          //     }
          //     return undefined;
          //   };

          //   // measuring point
          //   var mp = lookup(["Measuring Point","MeasuringPoint","Measuring point","Measuring_Point","Equipment"]);
          //   if (mp !== undefined) normalized.MeasuringPoint = String(mp).trim();

          //   // Reading / Counter column (accept many variants)
          //   var counterVal = lookup(["Reading","MeasurementCounterReading","Counter","Counter/Reading","MeasurementReading","Value"]);
          //   if (counterVal !== undefined) {
          //     var n = Number(counterVal);
          //     if (isNaN(n)) {
          //       var cleaned = String(counterVal).replace(/[^0-9\.\-\,]/g, "").replace(",", ".");
          //       n = Number(cleaned);
          //     }
          //     normalized.MeasurementCounterReading = !isNaN(n) ? n : counterVal;
          //     normalized.MeasurementReading = normalized.MeasurementCounterReading;
          //     normalized.Counter = normalized.MeasurementCounterReading;
          //   }

          //   // Difference column (accept "Difference" or API field name variants)
          // // inside your aResults.map normalization block
          // // accept both "Difference" and some legacy names
          // var diffVal = lookup(["Difference","MsmtCounterReadingDifference","MsmtCntrReadingDifference","DifferenceValue"]);
          // if (diffVal !== undefined) {
          //   var dnum = Number(diffVal);
          //   if (isNaN(dnum)) {
          //     var cleaned = String(diffVal).replace(/[^0-9\.\-\,]/g,"").replace(",",".");
          //     dnum = Number(cleaned);
          //   }
          //   normalized.MeasurementCounterReadingDifference = !isNaN(dnum) ? dnum : diffVal;
          // }

          //   // Read By column (accept "Read By" and "Ready By" variants)
          //   var readByVal = lookup(["Read By","ReadBy","Ready By","ReadyBy","MsmtRdngByUser","User"]);
          //   if (readByVal !== undefined) normalized.MsmtRdngByUser = String(readByVal).trim();

          //   // other fields...
          //   normalized.MeasurementDocumentText = lookup(["MeasurementDocumentText","Text","LongText","note"]) || "";
          //   normalized.PostingDate = lookup(["Posting Date","MsmtRdngDate","Date","PostingDate"]) || "";

          //   // return normalized row (ensures required properties set if present)
          //   return normalized;
          // });
          // inside reader.onload -> aResults mapping
          aResults = aResults.map(function (row) {
            var normalized = {};
            Object.keys(row || {}).forEach(function (k) { normalized[k] = row[k]; });

            var lookup = function (names) {
              for (var i = 0; i < names.length; i++) {
                var n = names[i];
                if (Object.prototype.hasOwnProperty.call(row, n) && row[n] !== undefined && row[n] !== null && String(row[n]).trim() !== "") {
                  return row[n];
                }
              }
              return undefined;
            };

            // Measuring point
            var mp = lookup(["Measuring Point", "MeasuringPoint", "Measuring point", "Measuring_Point", "Equipment"]);
            if (mp !== undefined) normalized.MeasuringPoint = String(mp).trim();

            // Reading / Counter column (optional)
            var readingVal = lookup(["Reading", "MeasurementReading", "MeasurementCounterReading", "Counter", "Value"]);
            if (readingVal !== undefined) {
              // try numeric conversion
              var n = Number(readingVal);
              if (isNaN(n)) {
                var cleaned = String(readingVal).replace(/[^0-9\.\-\,]/g, "").replace(",", ".");
                n = Number(cleaned);
              }
              normalized.MeasurementReading = !isNaN(n) ? n : readingVal;
              normalized.Counter = normalized.MeasurementReading;
            }

            // Difference column (optional)
            var diffVal = lookup(["Difference", "MsmtCounterReadingDifference", "CounterDifference"]);
            if (diffVal !== undefined) {
              var nd = Number(diffVal);
              if (isNaN(nd)) {
                var cleaned2 = String(diffVal).replace(/[^0-9\.\-\,]/g, "").replace(",", ".");
                nd = Number(cleaned2);
              }
              normalized.MsmtCounterReadingDifference = !isNaN(nd) ? nd : diffVal;
              // set flag to mark difference entered
              normalized.MsmtCntrReadingDiffIsEntered = true;
            } else {
              normalized.MsmtCntrReadingDiffIsEntered = false;
            }

            // Read/Ready by
            var readByVal = lookup(["Read By", "ReadBy", "Ready By", "ReadyBy", "MsmtRdngByUser", "User"]);
            if (readByVal !== undefined) normalized.MsmtRdngByUser = String(readByVal).trim();

            // Posting date - keep raw + normalized for UI
            var pd = lookup(["Posting Date", "MsmtRdngDate", "Date", "PostingDate",
              "Posting Date (DD-MM-YYYY)", "Posting Date (MM-DD-YYYY)", "Posting Date(DD-MM-YYYY)", "PostingDate(DD-MM-YYYY)"]);
            normalized.PostingDateRaw = pd; // keep original
            // convert to display format dd/mm/yyyy for UI
            // use helper _toYYYYMMDD if available — ensure it's called on controller instance later
            /*normalized.PostingDate = pd && that._toYYYYMMDD ? (function () {
              var ymd = that._toYYYYMMDD(pd); // returns YYYYMMDD or null
              if (!ymd) return "";
              //return ymd.slice(6, 8) + "/" + ymd.slice(4, 6) + "/" + ymd.slice(0, 4);
              return ymd;
            })() : (pd || "");*/

            //normalized.PostingDate = that._toYYYYMMDD(pd);
            normalized.PostingDate = that._convertDisplayDate(pd);

            // Text
            normalized.MeasurementDocumentText = lookup(["MeasurementDocumentText", "Text", "LongText", "note"]) || "";

            return normalized;
          });

          // optional: remove totally empty rows
          aResults = aResults.filter(function (r) {
            return r && (r.MeasuringPoint || r.MeasurementReading || r.MeasurementDocumentText);
          });

          // ENRICH: call your helper which reads OData for each measuring point
          // (this updates aResults in-place and refreshes the model)
          await that._enrichRowsWithMP(aResults);

          // put enriched rows into batchModel
          var batchModel = that.getOwnerComponent().getModel("batchModel");
          if (!batchModel) {
            batchModel = new sap.ui.model.json.JSONModel({ aEmployees: aResults });
            that.getOwnerComponent().setModel(batchModel, "batchModel");
          } else {
            batchModel.setProperty("/aEmployees", aResults);
            batchModel.refresh(true);
          }
        } catch (ex) {
          console.error("Import/enrich failed:", ex);
          sap.m.MessageBox.error("Import failed: " + (ex && ex.message ? ex.message : String(ex)));
        } finally {
          sap.ui.core.BusyIndicator.hide();  // <<< always hide busy
        }
      };

      reader.onerror = function (err) {
        console.error("File read error:", err);
        sap.m.MessageBox.error("Failed to read file");
      };

      reader.readAsBinaryString(file);
    },
    // add to your controller (near other helpers)
    _ensureXLSX: function () {
      var that = this;
      return new Promise(function (resolve, reject) {
        if (window.XLSX && window.XLSX.utils && typeof window.XLSX.utils.json_to_sheet === "function") {
          return resolve(window.XLSX);
        }
        // load SheetJS from CDN
        jQuery.getScript("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js")
          .done(function () {
            if (window.XLSX && window.XLSX.utils && typeof window.XLSX.utils.json_to_sheet === "function") {
              resolve(window.XLSX);
            } else {
              reject(new Error("XLSX loaded but utils.json_to_sheet not available"));
            }
          })
          .fail(function (err) { reject(err); });
      });
    },
    onPressDownload: function () {

      var that = this;
      this._ensureXLSX().then(function (XLSX) {
        // existing logic that uses XLSX.utils.json_to_sheet
        var excelColumnList = [{
          "Measuring Point": "", "Reading": "", "Difference": "", "Posting Date (DD-MM-YYYY)": "", "Text": "", "Read By": ""
        }];
        var ws = XLSX.utils.json_to_sheet(excelColumnList);
        ws["!cols"] = [{ width: 15 }, { width: 10 }, { width: 10 }, { width: 30 }, { width: 40 }, { width: 10 }];
        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
        XLSX.writeFile(wb, "Template.xlsx", { cellStyles: true });
        sap.m.MessageToast.show("Template File Downloading...");
      }).catch(function (err) {
        console.error("Failed to load XLSX:", err);
        sap.m.MessageBox.error("Failed to load Excel library. Please try again.");
      });
    },
    onClearLogs: function () {
      this.getOwnerComponent().getModel("ErrorListModel").setData({ errItems: [] });
      var oVBox = this.byId("idTableLayout");
      if (oVBox) { oVBox.setVisible(false); }
      try {
        this.byId("idTableLayout") && this.byId("idTableLayout").setVisible(false);
        this.byId("tblInput") && this.byId("tblInput").setVisible(true);
      } catch (e) { }
    },


    _convertDisplayDate(input) {
      if (!input) return "Invalid Format use DD-MM-YYYY";
      var s = String(input).trim();

      // Excel serial number (e.g., 45921 → 22/09/2025)
      if (/^\d{1,7}$/.test(s)) {
        var n = Number(s);
        if (n > 59) n -= 1; // 
        var epoch = new Date(Date.UTC(1899, 11, 30));
        epoch.setUTCDate(epoch.getUTCDate() + n);
        return String(epoch.getUTCDate()).padStart(2, "0") + "-" +
          String(epoch.getUTCMonth() + 1).padStart(2, "0") + "-" +
          epoch.getUTCFullYear();
      }

      // YYYYMMDD
      if (/^\d{8}$/.test(s)) {
        var y = s.substring(4, 8);
        var m = s.substring(2, 4);
        var d = s.substring(0, 2);
        return d + "-" + m + "-" + y;
      }

      // YYYY-MM-DD
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        var parts1 = s.split("-");
        return parts1[2] + "-" + parts1[1] + "-" + parts1[0];
      }

      // DD-MM-YYYY or DD/MM/YYYY
      if (/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(s)) {
        var parts2 = s.split(/[-/]/);
        return parts2[0] + "-" + parts2[1] + "-" + parts2[2];
      }

      // DD.MM.YYYY
      if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) {
        var parts3 = s.split(".");
        return parts3[0] + "-" + parts3[1] + "-" + parts3[2];
      }

      // DDMMYYYY (no delimiters, ensure not YYYYMMDD)
      if (/^\d{8}$/.test(s)) {
        var d2 = s.substring(0, 2);
        var m2 = s.substring(2, 4);
        var y2 = s.substring(4, 8);
        if (parseInt(d2, 10) <= 31 && parseInt(m2, 10) <= 12) {
          return d2 + "-" + m2 + "-" + y2;
        }
      }

      // SAP OData style: /Date(....)/
      if (s.indexOf("/Date(") === 0) {
        var ms = parseInt(s.replace("/Date(", "").replace(")/", ""), 10);
        var d3 = new Date(ms);
        return String(d3.getDate()).padStart(2, "0") + "-" +
          String(d3.getMonth() + 1).padStart(2, "0") + "-" +
          d3.getFullYear();
      }

      // Generic parse
      var d4 = new Date(s);
      if (!isNaN(d4.getTime())) {
        return String(d4.getDate()).padStart(2, "0") + "-" +
          String(d4.getMonth() + 1).padStart(2, "0") + "-" +
          d4.getFullYear();
      }
      
      return "Invalid Format use DD-MM-YYYY";
    }
  });
});