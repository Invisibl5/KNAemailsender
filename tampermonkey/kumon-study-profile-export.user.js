// ==UserScript==
// @name         Kumon - Study Profile Export (CSV)
// @namespace    kumon-automation
// @version      2.1
// @description  Export students + study profile. Batch fetch lowest PLANNED (uncompleted) page (min WorksheetNOFrom) per student+subject. Includes debug copy + auto-pagination for GetCenterAllStudentList (fix 100-student cap).
// @author       You
// @match        https://class-navi.digital.kumon.com/*
// @match        https://instructor2.digital.kumon.com/*
// @grant        none
// @run-at       document-start
// ==/UserScript==
//
// NOTE: This file is intentionally NOT part of Google Apps Script sync (clasp).

(function() {
  'use strict';

  var log = function() {
    try {
      if (typeof console !== 'undefined' && console.log) console.log.apply(console, ['[Kumon Study Profile]'].concat(Array.prototype.slice.call(arguments)));
    } catch (_) {}
  };

  var capturedStudents = [];
  var lastStudyResult = null;
  var lastStudyResultRequest = null;
  var lastStudyResultBySubject = {};
  var lastToken = null;
  var lastRegisterStudySetRequest = null;
  var lastRegisterStudySetResponse = null;
  var allStudyPlansData = { entries: [], fetchedAt: null };

  // NEW: lowest-page batch data
  var allLowestPagesData = { entries: [], fetchedAt: null };

  // NEW: student list capture meta (helps diagnose pagination / 100-cap)
  var lastStudentListMeta = { url: null, requestBody: null, receivedCount: 0, totalCount: null, capturedAt: null };

  var debugLogLines = [];
  var DEBUG_LOG_MAX = 300;

  function debugLog(msg) {
    var line = '[' + new Date().toLocaleTimeString() + '] ' + msg;
    debugLogLines.push(line);
    if (debugLogLines.length > DEBUG_LOG_MAX) debugLogLines.shift();
    log(msg);
    var el = document.getElementById('kumon-sp-debug-log');
    if (el) el.textContent = debugLogLines.join('\n');
  }

  function safeJsonParse(s) {
    if (!s || typeof s !== 'string') return null;
    try { return JSON.parse(s); } catch (e) { return null; }
  }

  function extractListFromResponse(data) {
    if (!data) return [];
    if (Array.isArray(data)) return data;
    if (data.StudentInfoList && Array.isArray(data.StudentInfoList)) return data.StudentInfoList;
    if (data.CenterAllStudentList && Array.isArray(data.CenterAllStudentList)) return data.CenterAllStudentList;
    if (data.StudentList && Array.isArray(data.StudentList)) return data.StudentList;
    if (data.students && Array.isArray(data.students)) return data.students;
    var first = Object.values(data).find(Array.isArray);
    return first || [];
  }

  function getCurrentTargetStudentLabel() {
    var ctx = getTargetContext();
    if (!ctx) return '(none – open a student’s Set screen to set target)';
    var sid = ctx.StudentID || '';
    var subj = ctx.SubjectCD === '010' ? 'Math' : ctx.SubjectCD === '022' ? 'Reading' : (ctx.SubjectCD || '');
    var ws = ctx.WorksheetCD || '';
    var cid = ctx.ClassID || '';
    var seq = ctx.ClassStudentSeq != null ? ctx.ClassStudentSeq : '';
    var name = '';
    if (sid && capturedStudents.length) {
      var s = capturedStudents.find(function(st) { return (st.StudentID || st.LoginID) === sid; });
      if (s) name = ' – ' + (s.FullName || s.StudentName || s.Name || '');
    }
    return 'StudentID ' + sid + name + ' | Subject ' + subj + ' | Worksheet ' + ws + ' | ClassID ' + cid + ' Seq ' + seq;
  }

  function getTargetContext() {
    if (lastStudyResultRequest) {
      return typeof lastStudyResultRequest === 'string'
        ? (function() { try { return JSON.parse(lastStudyResultRequest); } catch (e) { return null; } })()
        : lastStudyResultRequest;
    }
    return null;
  }

  // ---------- Inject capture into PAGE context ----------
  var INJECT_SCRIPT = function() {
    var isList = function(url) {
      var u = String(url);
      return u.indexOf('GetCenterAllStudentList') !== -1 || u.indexOf('StudentList') !== -1 || u.indexOf('GetStudentInfo') !== -1;
    };
    var isStudyResult = function(url) {
      return String(url).indexOf('GetStudyResultInfoList') !== -1;
    };
    var isRegisterStudySet = function(url) {
      return String(url).indexOf('RegisterStudySetInfo') !== -1;
    };
    var dispatchToken = function(authHeader) {
      if (!authHeader) return;
      try {
        document.dispatchEvent(new CustomEvent('KumonTokenCapture', { detail: { authorization: authHeader } }));
      } catch (e) {}
    };
    var dispatchRegisterStudySet = function(requestBody, responseData, authHeader) {
      try {
        document.dispatchEvent(new CustomEvent('KumonRegisterStudySetCapture', {
          detail: {
            requestBody: requestBody || null,
            responseJson: responseData ? JSON.stringify(responseData) : null,
            authorization: authHeader || null
          }
        }));
      } catch (e) {}
    };
    var origSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
    XMLHttpRequest.prototype.setRequestHeader = function(name, value) {
      if (name === 'Authorization' && value) {
        this._kumonAuth = value;
        dispatchToken(value);
      }
      return origSetRequestHeader.apply(this, arguments);
    };
    var extractList = function(data) {
      if (!data) return [];
      if (Array.isArray(data)) return data;
      if (data.StudentInfoList && Array.isArray(data.StudentInfoList)) return data.StudentInfoList;
      if (data.CenterAllStudentList && Array.isArray(data.CenterAllStudentList)) return data.CenterAllStudentList;
      if (data.StudentList && Array.isArray(data.StudentList)) return data.StudentList;
      if (data.students && Array.isArray(data.students)) return data.students;
      var first = Object.values(data).find(Array.isArray);
      return first || [];
    };
    var tryFindTotalCount = function(data) {
      if (!data || typeof data !== 'object') return null;
      var candidates = ['TotalCount', 'totalCount', 'total', 'Total', 'Count', 'count'];
      for (var i = 0; i < candidates.length; i++) {
        var k = candidates[i];
        if (data[k] != null && !isNaN(Number(data[k]))) return Number(data[k]);
      }
      // Sometimes nested
      if (data.Result && data.Result.TotalCount != null && !isNaN(Number(data.Result.TotalCount))) return Number(data.Result.TotalCount);
      return null;
    };

    var dispatchList = function(list, url, requestBody, rawData) {
      if (list.length === 0) return;
      try {
        document.dispatchEvent(new CustomEvent('KumonStudyProfileCapture', {
          detail: {
            studentsJson: JSON.stringify(list),
            url: url || null,
            requestBody: requestBody || null,
            totalCount: tryFindTotalCount(rawData),
            capturedAt: new Date().toISOString()
          }
        }));
      } catch (e) {}
    };
    var dispatchStudyResult = function(data, requestBody) {
      if (!data || typeof data !== 'object') return;
      try {
        document.dispatchEvent(new CustomEvent('KumonStudyResultCapture', {
          detail: { studyResultJson: JSON.stringify(data), requestBody: requestBody || null }
        }));
      } catch (e) {}
    };
    var origOpen = XMLHttpRequest.prototype.open;
    var origSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function(method, url) { this._kumonSpUrl = url; return origOpen.apply(this, arguments); };
    XMLHttpRequest.prototype.send = function(body) {
      var xhr = this;
      var url = xhr._kumonSpUrl || '';
      var reqBody = typeof body === 'string' ? body : null;
      if (isList(url)) {
        var done = false;
        xhr.addEventListener('readystatechange', function() {
          if (xhr.readyState !== 4 || done) return;
          done = true;
          try {
            var data = null;
            if (xhr.response != null && typeof xhr.response === 'object') data = xhr.response;
            else if (xhr.responseText) data = JSON.parse(xhr.responseText);
            if (data) dispatchList(extractList(data), url, reqBody, data);
          } catch (e) {}
        });
      }
      if (isStudyResult(url)) {
        var doneSr = false;
        xhr.addEventListener('readystatechange', function() {
          if (xhr.readyState !== 4 || doneSr) return;
          doneSr = true;
          try {
            var data = null;
            if (xhr.response != null && typeof xhr.response === 'object') data = xhr.response;
            else if (xhr.responseText) data = JSON.parse(xhr.responseText);
            if (data) dispatchStudyResult(data, reqBody);
          } catch (e) {}
        });
      }
      if (isRegisterStudySet(url)) {
        var doneReg = false;
        var auth = xhr._kumonAuth || null;
        xhr.addEventListener('readystatechange', function() {
          if (xhr.readyState !== 4 || doneReg) return;
          doneReg = true;
          try {
            var data = null;
            if (xhr.response != null && typeof xhr.response === 'object') data = xhr.response;
            else if (xhr.responseText) data = JSON.parse(xhr.responseText);
            dispatchRegisterStudySet(reqBody, data, auth);
          } catch (e) {}
        });
      }
      return origSend.apply(this, arguments);
    };
    var origFetch = window.fetch;
    window.fetch = function(input, init) {
      var url = typeof input === 'string' ? input : (input && input.url) || '';
      var reqBody = (init && init.body) ? (typeof init.body === 'string' ? init.body : null) : null;
      if (isList(url)) {
        return origFetch.apply(this, arguments).then(function(res) {
          var clone = res.clone();
          clone.json().then(function(data) { if (data) dispatchList(extractList(data), url, reqBody, data); }).catch(function() {});
          return res;
        });
      }
      if (isStudyResult(url)) {
        return origFetch.apply(this, arguments).then(function(res) {
          var clone = res.clone();
          clone.json().then(function(data) { if (data) dispatchStudyResult(data, reqBody); }).catch(function() {});
          return res;
        });
      }
      if (isRegisterStudySet(url)) {
        var authHeader = (init && init.headers && init.headers.Authorization) ? init.headers.Authorization : null;
        if (authHeader) dispatchToken(authHeader);
        return origFetch.apply(this, arguments).then(function(res) {
          var clone = res.clone();
          clone.json().then(function(data) { dispatchRegisterStudySet(reqBody, data, authHeader); }).catch(function() {});
          return res;
        });
      }
      return origFetch.apply(this, arguments);
    };
  };

  function injectPageScript() {
    if (window.__kumonStudyProfilePageScriptInjected) return;
    var script = document.createElement('script');
    script.textContent = '(' + INJECT_SCRIPT.toString() + ')();';
    var target = document.documentElement || document.head || document.body;
    if (target) {
      window.__kumonStudyProfilePageScriptInjected = true;
      target.appendChild(script);
      script.remove();
      log('page script injected (capture GetCenterAllStudentList)');
    } else {
      document.addEventListener('DOMContentLoaded', function runOnce() {
        document.removeEventListener('DOMContentLoaded', runOnce);
        injectPageScript();
      });
    }
  }
  injectPageScript();

  document.addEventListener('KumonStudyProfileCapture', function(ev) {
    var detail = ev.detail;
    if (!detail || !detail.studentsJson) return;
    try {
      var list = JSON.parse(detail.studentsJson);
      if (Array.isArray(list) && list.length > 0) {
        capturedStudents = list;
        lastStudentListMeta.url = detail.url || null;
        lastStudentListMeta.requestBody = detail.requestBody || null;
        lastStudentListMeta.receivedCount = list.length;
        lastStudentListMeta.totalCount = detail.totalCount != null ? detail.totalCount : null;
        lastStudentListMeta.capturedAt = detail.capturedAt || new Date().toISOString();
        log('captured', list.length, 'students for study profile');
        if (lastStudentListMeta.totalCount && lastStudentListMeta.totalCount > list.length) {
          debugLog('Student list looks paginated: got ' + list.length + ' / total ' + lastStudentListMeta.totalCount + '. Likely API returns first page (often 100).');
          debugLog('List URL: ' + (lastStudentListMeta.url || '(unknown)'));
        } else {
          debugLog('Student list captured: ' + list.length + (lastStudentListMeta.totalCount ? (' / total ' + lastStudentListMeta.totalCount) : ''));
        }
        updateStudyProfileUI();
      }
    } catch (e) {
      log('parse error', e);
    }
  });

  document.addEventListener('KumonStudyResultCapture', function(ev) {
    var detail = ev.detail;
    if (!detail || !detail.studyResultJson) return;
    try {
      lastStudyResult = JSON.parse(detail.studyResultJson);
      lastStudyResultRequest = detail.requestBody || null;
      var ctx = typeof lastStudyResultRequest === 'string'
        ? (function() { try { return JSON.parse(lastStudyResultRequest); } catch (e) { return null; } })()
        : lastStudyResultRequest;
      var subjectCD = (ctx && ctx.SubjectCD) ? String(ctx.SubjectCD) : '';
      if (subjectCD) {
        lastStudyResultBySubject[subjectCD] = { result: lastStudyResult, request: lastStudyResultRequest };
      }
      var units = (lastStudyResult.StudyUnitInfoList && lastStudyResult.StudyUnitInfoList.length) || 0;
      var sid = ctx ? (ctx.StudentID || '') : '';
      var subj = ctx ? (ctx.SubjectCD === '010' ? 'Math' : ctx.SubjectCD === '022' ? 'Reading' : ctx.SubjectCD) : '';
      debugLog('GetStudyResultInfoList captured: ' + units + ' units | target StudentID=' + sid + ' Subject=' + subj);
      updateStudyProfileUI();
    } catch (e) {
      debugLog('Study result parse error: ' + (e && e.message));
    }
  });

  document.addEventListener('KumonTokenCapture', function(ev) {
    var auth = ev.detail && ev.detail.authorization;
    if (auth) {
      lastToken = auth;
      debugLog('Token captured (length=' + (auth ? auth.length : 0) + ')');
      updateStudyProfileUI();
    }
  });

  document.addEventListener('KumonRegisterStudySetCapture', function(ev) {
    var d = ev.detail;
    if (!d) return;
    try {
      if (d.requestBody) lastRegisterStudySetRequest = typeof d.requestBody === 'string' ? JSON.parse(d.requestBody) : d.requestBody;
      if (d.responseJson) lastRegisterStudySetResponse = JSON.parse(d.responseJson);
      if (d.authorization) lastToken = d.authorization;
      updateStudyProfileUI();
    } catch (e) {
      debugLog('RegisterStudySet capture parse error: ' + (e && e.message));
    }
  });

  function subjectName(subjectCD) {
    var cd = String(subjectCD || '');
    if (cd === '010') return 'Math';
    if (cd === '022') return 'Reading';
    return cd ? 'Subject' + cd : '';
  }

  function isStudyUnitCompleted(u) {
    if (!u) return false;
    if (u.StudyStatus === '6') return true;
    if (u.StudyDate || u.FinishDate) return true;
    return false;
  }

  function splitCompletedAndPlanned(list) {
    var completed = [];
    var planned = [];
    (list || []).forEach(function(u) {
      if (isStudyUnitCompleted(u)) completed.push(u);
      else planned.push(u);
    });
    return { completed: completed, planned: planned };
  }

  /** Compute the lowest page range in PLANNED (uncompleted) StudyUnitInfoList (min WorksheetNOFrom). */
  function computeLowestPages(studyResult) {
    var list = (studyResult && studyResult.StudyUnitInfoList) ? studyResult.StudyUnitInfoList : [];
    // Only planned/uncompleted work (ignore history)
    list = splitCompletedAndPlanned(list).planned || [];
    var minFrom = null;
    var minTo = null;
    var minRow = null;
    for (var i = 0; i < list.length; i++) {
      var u = list[i];
      if (u == null) continue;
      var from = u.WorksheetNOFrom;
      var to = u.WorksheetNOTo;
      if (from == null || from === '') continue;
      var fromN = Number(from);
      var toN = (to == null || to === '') ? null : Number(to);
      if (isNaN(fromN)) continue;
      if (minFrom === null || fromN < minFrom) {
        minFrom = fromN;
        minTo = (toN != null && !isNaN(toN)) ? toN : null;
        minRow = (u.StudyScheduleIndex != null) ? u.StudyScheduleIndex : null;
      }
    }
    return { minFrom: minFrom, minTo: minTo, row: minRow };
  }

  // ---------- Batch fetch helpers ----------
  var GET_STUDY_RESULT_URL = 'https://instructor2.digital.kumon.com/USA/api/ATD0010P/GetStudyResultInfoList';

  function buildGetStudyResultRequestBody(student, study, baseContext) {
    var base = typeof baseContext === 'string'
      ? (function() { try { return JSON.parse(baseContext); } catch (e) { return null; } })()
      : baseContext;
    if (!base) return null;
    return {
      CenterID: base.CenterID || '',
      ClassID: (study && study.ClassID) != null ? String(study.ClassID) : (base.ClassID || ''),
      ClassStudentSeq: study && study.ClassStudentSeq != null ? study.ClassStudentSeq : base.ClassStudentSeq,
      StudentID: (student && (student.StudentID || student.LoginID)) || base.StudentID || '',
      SubjectCD: (study && study.SubjectCD) != null ? String(study.SubjectCD) : (base.SubjectCD || ''),
      SystemCountryCD: base.SystemCountryCD || 'USA',
      WorksheetCD: (study && study.NextWorksheetCD) ? String(study.NextWorksheetCD) : (base.WorksheetCD || '')
    };
  }

  function getStudyPairsQueue() {
    var queue = [];
    capturedStudents.forEach(function(student) {
      var list = student.StudentStudyInfoList || student.StudyInfoList || [];
      var fullName = student.FullName || student.StudentName || student.Name || '';
      var sid = student.StudentID || student.LoginID || '';
      list.forEach(function(study) {
        if (!study.SubjectCD || !study.NextWorksheetCD) return;
        queue.push({ student: student, study: study, fullName: fullName, sid: sid });
      });
    });
    return queue;
  }

  // ---------- NEW: auto-pagination for student list (fix 100 cap) ----------
  function getStudentKey(st) {
    return String((st && (st.StudentID || st.LoginID || st.StudentId || st.loginId)) || '');
  }

  function mergeUniqueStudents(existing, incoming) {
    var map = {};
    var out = [];
    (existing || []).forEach(function(s) {
      var k = getStudentKey(s);
      if (!k) return;
      if (!map[k]) { map[k] = true; out.push(s); }
    });
    (incoming || []).forEach(function(s) {
      var k = getStudentKey(s);
      if (!k) return;
      if (!map[k]) { map[k] = true; out.push(s); }
    });
    return out;
  }

  /**
   * Attempts to fetch ALL students by paging through GetCenterAllStudentList.
   * Kumon often returns first 100 only; this tries common pagination fields.
   */
  function fetchAllStudentsPaginated(options, callback) {
    var onProgress = (options && options.onProgress) || function() {};
    var onComplete = (options && options.onComplete) || callback || function() {};
    var pageSize = (options && options.pageSize) || 100;
    var maxPages = (options && options.maxPages) || 20; // 20*100 = 2000 safety cap

    if (!lastToken) {
      onProgress('No token. Open any student’s Set screen first.');
      onComplete('No token', null);
      return;
    }
    if (!lastStudentListMeta.url) {
      onProgress('No list URL captured yet. Open the student list page first.');
      onComplete('No list URL', null);
      return;
    }

    var baseBody = safeJsonParse(lastStudentListMeta.requestBody) || {};
    var url = lastStudentListMeta.url;

    // Try a few common schemas. We'll pick the first one that yields >0 on page 2.
    var schemas = [
      function(offset, pageNo) { var b = Object.assign({}, baseBody); b.StartNum = offset; b.DispNum = pageSize; return b; },
      function(offset, pageNo) { var b = Object.assign({}, baseBody); b.StartIndex = offset; b.Count = pageSize; return b; },
      function(offset, pageNo) { var b = Object.assign({}, baseBody); b.Offset = offset; b.Limit = pageSize; return b; },
      function(offset, pageNo) { var b = Object.assign({}, baseBody); b.PageNo = pageNo; b.DispNum = pageSize; return b; },
      function(offset, pageNo) { var b = Object.assign({}, baseBody); b.Page = pageNo; b.PageSize = pageSize; return b; }
    ];

    function doFetch(body) {
      return fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': lastToken },
        body: JSON.stringify(body)
      }).then(function(res) { return res.json(); });
    }

    var chosenSchemaIdx = null;
    var all = [];

    // Always start from offset 0
    onProgress('Fetching page 1 …');
    doFetch(schemas[0](0, 1)).then(function(data1) {
      var list1 = extractListFromResponse(data1);
      all = mergeUniqueStudents(all, list1);
      onProgress('Page 1: got ' + list1.length + ', unique=' + all.length);

      // If it already returned more than 100, we're done.
      if (list1.length < pageSize) {
        capturedStudents = all;
        updateStudyProfileUI();
        onComplete(null, { students: all, pages: 1, schema: 'single' });
        return;
      }

      // Probe page 2 with each schema to find one that works
      var probeOffset = pageSize;
      var probePageNo = 2;

      var probePromises = schemas.map(function(makeBody, idx) {
        return doFetch(makeBody(probeOffset, probePageNo))
          .then(function(d) { return { idx: idx, data: d, ok: true }; })
          .catch(function(e) { return { idx: idx, data: null, ok: false, err: e }; });
      });

      return Promise.all(probePromises).then(function(results) {
        // Pick schema that returns non-empty list and adds new IDs
        for (var i = 0; i < results.length; i++) {
          var r = results[i];
          if (!r.ok || !r.data) continue;
          var list2 = extractListFromResponse(r.data);
          if (!list2 || list2.length === 0) continue;
          var merged = mergeUniqueStudents(all, list2);
          if (merged.length > all.length) {
            chosenSchemaIdx = r.idx;
            all = merged;
            onProgress('Pagination schema #' + chosenSchemaIdx + ' works (page 2: +' + (merged.length - (merged.length - list2.length)) + ').');
            break;
          }
        }

        if (chosenSchemaIdx == null) {
          onProgress('Could not find a working pagination schema. Still only have ' + all.length + '.');
          capturedStudents = all;
          updateStudyProfileUI();
          onComplete(null, { students: all, pages: 1, schema: null });
          return;
        }

        // Continue paging using chosen schema
        var pageNo = 2;
        var offset = pageSize;
        var pagesFetched = 2;

        function nextPage() {
          pageNo += 1;
          offset += pageSize;
          pagesFetched += 1;
          if (pagesFetched > maxPages) {
            onProgress('Stopped at maxPages=' + maxPages + '. unique=' + all.length);
            capturedStudents = all;
            updateStudyProfileUI();
            onComplete(null, { students: all, pages: pagesFetched - 1, schema: chosenSchemaIdx });
            return;
          }

          onProgress('Fetching page ' + pageNo + ' … unique=' + all.length);
          return doFetch(schemas[chosenSchemaIdx](offset, pageNo)).then(function(d) {
            var list = extractListFromResponse(d);
            var before = all.length;
            all = mergeUniqueStudents(all, list);
            var added = all.length - before;
            onProgress('Page ' + pageNo + ': got ' + list.length + ', added ' + added + ', unique=' + all.length);
            if (!list || list.length === 0 || added === 0) {
              capturedStudents = all;
              updateStudyProfileUI();
              onComplete(null, { students: all, pages: pageNo, schema: chosenSchemaIdx });
              return;
            }
            return new Promise(function(resolve) { setTimeout(resolve, 250); }).then(nextPage);
          }).catch(function(err) {
            onProgress('Error page ' + pageNo + ': ' + (err && err.message));
            capturedStudents = all;
            updateStudyProfileUI();
            onComplete(err, { students: all, pages: pageNo - 1, schema: chosenSchemaIdx });
          });
        }

        // We already merged a working page 2 in the probe loop above only if it added new IDs.
        // Keep going from page 3.
        return new Promise(function(resolve) { setTimeout(resolve, 250); }).then(nextPage);
      });
    }).catch(function(err) {
      onProgress('List fetch failed: ' + (err && err.message));
      onComplete(err, null);
    });
  }

  // NEW: Fetch all lowest pages across all students/subjects
  function fetchAllLowestPages(options, callback) {
    var testMode = options && options.testMode;
    var onProgress = (options && options.onProgress) || function() {};
    var onComplete = (options && options.onComplete) || callback || function() {};

    if (!lastToken) {
      onProgress('No token. Open any student’s Set screen first.');
      onComplete('No token', null);
      return;
    }
    var baseContext = lastStudyResultRequest;
    var req = typeof baseContext === 'string'
      ? (function() { try { return JSON.parse(baseContext); } catch (e) { return null; } })()
      : baseContext;
    if (!req || !req.CenterID) {
      onProgress('No center context. Open any student’s Set screen once.');
      onComplete('No center context', null);
      return;
    }

    var queue = getStudyPairsQueue();
    if (queue.length === 0) {
      onProgress('No students/subjects. Load student list (with study info).');
      onComplete('No queue', null);
      return;
    }
    if (testMode) queue = queue.slice(0, 2);

    allLowestPagesData.entries = [];
    allLowestPagesData.fetchedAt = new Date().toISOString();

    var delayMs = (options && options.delayMs) || 400;
    var index = 0;

    function next() {
      if (index >= queue.length) {
        onProgress('Done. ' + allLowestPagesData.entries.length + ' rows.');
        onComplete(null, allLowestPagesData);
        return;
      }

      var item = queue[index];
      index += 1;
      onProgress('Fetching ' + index + '/' + queue.length + ' – ' + (item.fullName || item.sid) + ' …');

      var body = buildGetStudyResultRequestBody(item.student, item.study, baseContext);
      if (!body) { setTimeout(next, delayMs); return; }

      fetch(GET_STUDY_RESULT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': lastToken },
        body: JSON.stringify(body)
      })
        .then(function(res) { return res.json().then(function(data) { return { ok: res.ok, data: data }; }); })
        .then(function(r) {
          if (!r.ok || (r.data && r.data.Result && r.data.Result.ResultCode !== 0)) {
            setTimeout(next, delayMs);
            return;
          }
          var result = r.data;
          var lowest = computeLowestPages(result);
          var subj = subjectName(body.SubjectCD);
          allLowestPagesData.entries.push({
            studentId: item.sid,
            fullName: item.fullName,
            subject: subj,
            level: String(body.WorksheetCD || ''),
            lowestFrom: lowest.minFrom != null ? lowest.minFrom : '',
            lowestTo: lowest.minTo != null ? lowest.minTo : '',
            lowestIndex: lowest.row != null ? lowest.row : ''
          });
          setTimeout(next, delayMs);
        })
        .catch(function(err) {
          debugLog('Fetch error: ' + (err && err.message));
          setTimeout(next, delayMs);
        });
    }

    next();
  }

  function generateLowestPagesCSV() {
    if (!allLowestPagesData.entries || allLowestPagesData.entries.length === 0) return '';
    var rows = ['StudentID\tFullName\tSubject\tLevel\tLowestPlannedFrom\tLowestPlannedTo\tLowestPlannedIndex'];
    allLowestPagesData.entries.forEach(function(e) {
      rows.push([e.studentId, e.fullName, e.subject, e.level, e.lowestFrom, e.lowestTo, e.lowestIndex].join('\t'));
    });
    return rows.join('\n');
  }

  // ---------- UI ----------
  function updateStudyProfileUI() {
    var countEl = document.getElementById('kumon-study-profile-count');
    if (countEl) {
      var base = capturedStudents.length + ' student' + (capturedStudents.length === 1 ? '' : 's');
      if (lastStudentListMeta.totalCount) base += ' (total ' + lastStudentListMeta.totalCount + ')';
      countEl.textContent = base;
    }

    var srEl = document.getElementById('kumon-sp-study-result-status');
    if (srEl) {
      if (!lastStudyResult) srEl.textContent = 'Not captured. Open a student’s Set screen (セット画面).';
      else srEl.textContent = 'Captured study result for current student/subject.';
    }

    var pasteStatus = document.getElementById('kumon-sp-paste-status');
    if (pasteStatus) pasteStatus.textContent = lastToken ? 'Token: captured.' : 'Token: not captured. Open any student.';

    var targetEl = document.getElementById('kumon-sp-current-target');
    if (targetEl) targetEl.textContent = getCurrentTargetStudentLabel();
  }

  function injectStudyProfileUI() {
    if (document.getElementById('kumon-study-profile-panel')) return;

    var panel = document.createElement('div');
    panel.id = 'kumon-study-profile-panel';
    panel.innerHTML =
      '<div class="kumon-sp-head">Kumon Study Profile Export</div>' +
      '<div class="kumon-sp-body">' +
      '  <div class="kumon-sp-block">' +
      '    <div class="kumon-sp-label">Students (from list capture)</div>' +
      '    <div id="kumon-study-profile-count" class="kumon-sp-count">0 students</div>' +
      '    <div class="kumon-sp-hint">Load the student list page; data is captured automatically.</div>' +
      '  </div>' +
      '  <div class="kumon-sp-block">' +
      '    <button id="kumon-sp-fetch-all-students-btn" class="kumon-sp-btn-primary">Fetch ALL students (fix 100 cap)</button>' +
      '    <button id="kumon-sp-fetch-lowest-btn" class="kumon-sp-btn-primary">Fetch lowest pages (all students)</button>' +
      '    <button id="kumon-sp-download-lowest-btn" class="kumon-sp-btn-secondary">Download lowest pages (CSV)</button>' +
      '    <button id="kumon-sp-copy-debug-btn" class="kumon-sp-btn-secondary">Copy debug</button>' +
      '    <div id="kumon-sp-fetch-status" class="kumon-sp-hint" style="min-height:1.2em;"></div>' +
      '  </div>' +
      '  <div class="kumon-sp-block">' +
      '    <div class="kumon-sp-label">Current target</div>' +
      '    <div id="kumon-sp-current-target" class="kumon-sp-hint" style="font-weight:600;"></div>' +
      '    <div id="kumon-sp-paste-status" class="kumon-sp-hint"></div>' +
      '  </div>' +
      '  <details class="kumon-sp-details"><summary>Debug</summary>' +
      '    <pre id="kumon-sp-debug-log" class="kumon-sp-pre" style="max-height:220px;"></pre>' +
      '  </details>' +
      '</div>';

    var style = document.createElement('style');
    style.textContent =
      '#kumon-study-profile-panel{position:fixed;top:20px;right:20px;width:320px;min-width:260px;max-width:96vw;' +
      'display:flex;flex-direction:column;overflow:hidden;background:#1e1e2e;color:#cdd6f4;' +
      'font:13px/1.45 "Segoe UI",system-ui,sans-serif;border:1px solid rgba(137,180,250,0.25);border-radius:14px;z-index:2147483644;' +
      'box-shadow:0 16px 48px rgba(0,0,0,0.4);}' +
      '.kumon-sp-head{padding:12px 14px;font-weight:600;font-size:14px;background:rgba(137,180,250,0.1);border-bottom:1px solid rgba(137,180,250,0.2);}' +
      '.kumon-sp-body{flex:1;overflow:auto;padding:12px;}' +
      '.kumon-sp-block{margin-bottom:14px;}' +
      '.kumon-sp-label{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;color:#89b4fa;margin-bottom:4px;}' +
      '.kumon-sp-count{font-size:20px;font-weight:700;color:#89b4fa;}' +
      '.kumon-sp-hint{font-size:11px;color:#6c7086;margin-top:6px;line-height:1.4;}' +
      '.kumon-sp-btn-primary{width:100%;padding:10px 14px;margin-bottom:8px;background:linear-gradient(135deg,#89b4fa,#7c9ee0);color:#1e1e2e;border:none;border-radius:10px;font-weight:600;font-size:13px;cursor:pointer;}' +
      '.kumon-sp-btn-secondary{width:100%;padding:8px 12px;margin-bottom:8px;font-size:12px;border-radius:8px;border:1px solid rgba(137,180,250,0.35);background:rgba(49,50,68,0.5);color:#cdd6f4;cursor:pointer;}' +
      '.kumon-sp-details{margin-top:10px;border:1px solid rgba(137,180,250,0.15);border-radius:8px;overflow:hidden;}' +
      '.kumon-sp-details summary{padding:6px 10px;font-size:12px;font-weight:600;color:#a6adc8;cursor:pointer;background:rgba(0,0,0,0.2);}' +
      '.kumon-sp-pre{font-size:10px;background:rgba(0,0,0,0.35);padding:8px;margin:8px;border-radius:6px;overflow:auto;white-space:pre-wrap;word-break:break-word;}';

    document.documentElement.appendChild(style);
    document.body.appendChild(panel);

    var statusEl = document.getElementById('kumon-sp-fetch-status');

    document.getElementById('kumon-sp-fetch-lowest-btn').addEventListener('click', function() {
      if (statusEl) statusEl.textContent = 'Fetching…';
      fetchAllLowestPages({
        delayMs: 400,
        onProgress: function(msg) {
          if (statusEl) statusEl.textContent = msg;
          debugLog(msg);
        },
        onComplete: function(err, data) {
          if (err) {
            if (statusEl) statusEl.textContent = 'Error: ' + err;
            return;
          }
          var n = (data && data.entries && data.entries.length) || 0;
          if (statusEl) statusEl.textContent = 'Done. ' + n + ' rows. Click Download.';
        }
      });
    });

    document.getElementById('kumon-sp-download-lowest-btn').addEventListener('click', function() {
      var csv = generateLowestPagesCSV();
      if (!csv) { alert('No data. Click “Fetch lowest pages” first.'); return; }
      var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'kumon-lowest-pages-' + new Date().toISOString().slice(0, 10) + '.csv';
      a.click();
      URL.revokeObjectURL(a.href);
    });

    document.getElementById('kumon-sp-fetch-all-students-btn').addEventListener('click', function() {
      if (statusEl) statusEl.textContent = 'Fetching full student list…';
      fetchAllStudentsPaginated({
        pageSize: 100,
        maxPages: 20,
        onProgress: function(msg) {
          if (statusEl) statusEl.textContent = msg;
          debugLog(msg);
        },
        onComplete: function(err, data) {
          if (err) {
            if (statusEl) statusEl.textContent = 'Error: ' + err;
            return;
          }
          var n = (data && data.students && data.students.length) || capturedStudents.length;
          if (statusEl) statusEl.textContent = 'Students loaded: ' + n + '.';
          updateStudyProfileUI();
        }
      });
    });

    document.getElementById('kumon-sp-copy-debug-btn').addEventListener('click', function() {
      var meta = [
        'CapturedStudents=' + (capturedStudents ? capturedStudents.length : 0),
        'TotalCount=' + (lastStudentListMeta.totalCount != null ? lastStudentListMeta.totalCount : ''),
        'ListURL=' + (lastStudentListMeta.url || ''),
        'ListRequestBody=' + (lastStudentListMeta.requestBody || ''),
        'CapturedAt=' + (lastStudentListMeta.capturedAt || ''),
        '--- Debug ---'
      ].join('\n');
      var text = meta + '\n' + debugLogLines.join('\n');
      try {
        navigator.clipboard.writeText(text).then(function() {
          alert('Debug copied to clipboard.');
        }).catch(function() {
          prompt('Copy debug:', text);
        });
      } catch (e) {
        prompt('Copy debug:', text);
      }
    });

    debugLog('Panel ready. Load student list, then open any student Set screen once (to capture token).');
    updateStudyProfileUI();
  }

  if (document.body) injectStudyProfileUI();
  else document.addEventListener('DOMContentLoaded', injectStudyProfileUI);
})();

