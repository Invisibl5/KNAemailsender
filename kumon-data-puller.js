#!/usr/bin/env node
/**
 * Kumon Data Puller - Standalone script to authenticate and pull data
 *
 * Usage (interactive):
 *   node kumon-data-puller.js
 *
 * Usage (lowest pages CSV - terminal):
 *   node kumon-data-puller.js --lowest --list=students.txt [--subject=both|010|022] [--output=out.csv]
 *   Put one student per line in students.txt: LoginID<tab>Name (tab or comma, name ignored).
 *   Prompts for username + password hash (from NaviPasswordHash cookie), then fetches lowest
 *   planned worksheet per student/subject and prints CSV to stdout.
 */

const readline = require('readline');
const https = require('https');
const { URL } = require('url');

// Parse command line args
const args = process.argv.slice(2);
const config = {};
args.forEach(arg => {
  if (arg === '--lowest') {
    config.lowest = true;
  } else {
    const match = arg.match(/^--(\w+)=(.*)$/);
    if (match) config[match[1]] = match[2];
  }
});

const BASE_URL = 'https://instructor2.digital.kumon.com/USA';

// Client object for API calls
function clientObject(id) {
  return {
    applicationName: 'Class-Navi',
    version: '1.0.0.0',
    programName: 'Class-Navi',
    machineName: '-',
    os: process.platform + ' ' + process.arch,
    id: String(id != null ? id : Date.now())
  };
}

// HTTP request helper
function httpRequest(url, options = {}) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const headers = {
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36',
      ...options.headers
    };
    if (options.contentType) {
      headers['Content-Type'] = options.contentType;
    } else if (!headers['Content-Type']) {
      headers['Content-Type'] = 'application/json';
    }
    
    const opts = {
      hostname: urlObj.hostname,
      port: urlObj.port || 443,
      path: urlObj.pathname + urlObj.search,
      method: options.method || 'GET',
      headers: headers
    };
    
    const req = https.request(opts, (res) => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const json = JSON.parse(data);
            resolve({ status: res.statusCode, data: json, headers: res.headers });
          } catch (e) {
            resolve({ status: res.statusCode, data: data, headers: res.headers });
          }
        } else {
          resolve({ status: res.statusCode, data: data, headers: res.headers });
        }
      });
    });
    
    req.on('error', reject);
    if (options.body) req.write(options.body);
    req.end();
  });
}

// Login and get access token
async function login(username, password, isHash = false) {
  console.log('Logging in...');
  // Format: grant_type=password&username=USA%2F{LoginID}&password={passwordHash}
  // Password hash from NaviPasswordHash cookie: decode it first, then encode for form
  const usernameEncoded = encodeURIComponent(username.trim());
  let passwordEncoded;
  
  if (isHash) {
    // Hash from cookie is already URL-encoded (has %2F, %3D), use as-is in form body
    passwordEncoded = password.trim();
  } else {
    passwordEncoded = encodeURIComponent(password.trim());
  }
  
  const body = `grant_type=password&username=USA%2F${usernameEncoded}&password=${passwordEncoded}`;
  
  if (process.env.DEBUG) {
    console.log('Request body:', body.replace(/password=[^&]+/, 'password=***'));
  }
  
  const res = await httpRequest(BASE_URL + '/token', {
    method: 'POST',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      'Accept': 'application/json, text/plain, */*',
      'Accept-Language': 'en-US,en;q=0.9',
      'Origin': 'https://instructor2.digital.kumon.com',
      'Referer': 'https://instructor2.digital.kumon.com/USA/'
    },
    body: body
  });
  
  // Check if we got HTML error page
  if (typeof res.data === 'string' && res.data.includes('<!DOCTYPE html>')) {
    const errorMatch = res.data.match(/<H1>(.*?)<\/H1>/);
    const errorMsg = errorMatch ? errorMatch[1] : 'HTML error page';
    throw new Error('Login failed: Server returned HTML error page (' + res.status + ').\n' +
      'Error: ' + errorMsg + '\n' +
      'This might mean:\n' +
      '1. Password hash expired (get fresh hash from browser after logging in)\n' +
      '2. Server is blocking non-browser requests\n' +
      '3. Request format is incorrect\n\n' +
      'Try: Log into Kumon via browser, get fresh NaviPasswordHash cookie, then run script again.');
  }
  
  if (res.status !== 200) {
    const errorPreview = typeof res.data === 'string' 
      ? res.data.substring(0, 300) 
      : JSON.stringify(res.data).substring(0, 300);
    throw new Error('Login failed (status ' + res.status + '): ' + errorPreview);
  }
  
  if (!res.data || !res.data.access_token) {
    throw new Error('Login failed: No access_token in response. Response: ' + JSON.stringify(res.data).substring(0, 200));
  }
  
  const token = res.data.access_token;
  console.log('✓ Login successful. Token expires in', res.data.expires_in, 'seconds.');
  return { token, refreshToken: res.data.refresh_token, expiresIn: res.data.expires_in };
}

// API call helper
async function apiCall(token, endpoint, body) {
  const res = await httpRequest(BASE_URL + endpoint, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token },
    body: JSON.stringify(body)
  });
  
  if (res.status !== 200) {
    throw new Error(`API call failed (${res.status}): ${JSON.stringify(res.data)}`);
  }
  
  if (res.data.Result && res.data.Result.ResultCode !== 0) {
    const errors = res.data.Result.Errors || [];
    throw new Error(`API error: ${errors.map(e => e.Message || e.ErrorCode).join(', ') || 'ResultCode ' + res.data.Result.ResultCode}`);
  }
  
  return res.data;
}

// Get instructor info (to get CenterID, etc.)
async function getInstructorInfo(token, loginID) {
  return apiCall(token, '/api/ATX0010P/GetInstructorInfo', {
    SystemCountryCD: 'USA',
    LoginID: loginID,
    client: clientObject()
  });
}

// Extract student list from API response (multiple possible keys)
function extractStudentList(res) {
  if (!res) return [];
  if (Array.isArray(res)) return res;
  if (res.CenterAllStudentList && Array.isArray(res.CenterAllStudentList)) return res.CenterAllStudentList;
  if (res.StudentInfoList && Array.isArray(res.StudentInfoList)) return res.StudentInfoList;
  if (res.StudentList && Array.isArray(res.StudentList)) return res.StudentList;
  const first = Object.values(res).find(v => Array.isArray(v));
  return first || [];
}

// Get all students (paginated); try Offset/GetNum first, then StartNum/DispNum if first page empty
async function getAllStudents(token, centerID, instructorID, instructorAssistantSec) {
  const students = [];
  const pageSize = 100;
  const baseBody = {
    SystemCountryCD: 'USA',
    CenterID: centerID,
    InstructorID: instructorID,
    InstructorAssistantSec: instructorAssistantSec,
    ValidFlg: '1',
    client: clientObject()
  };

  let useStartNum = false;
  let startNum = 1;

  const res = await apiCall(token, '/api/ATE0010P/GetCenterAllStudentList', {
    ...baseBody,
    Offset: 1,
    GetNum: pageSize
  });
  let list = extractStudentList(res);
  if (list.length === 0) {
    const res2 = await apiCall(token, '/api/ATE0010P/GetCenterAllStudentList', {
      ...baseBody,
      StartNum: 1,
      DispNum: pageSize
    });
    list = extractStudentList(res2);
    useStartNum = true;
  }

  while (list.length > 0) {
    students.push(...list);
    if (process.stderr.isTTY) console.error(`  Fetched ${students.length} students so far...`);
    if (list.length < pageSize) break;
    await new Promise(r => setTimeout(r, 500));
    if (useStartNum) {
      startNum = 1 + students.length;
      const next = await apiCall(token, '/api/ATE0010P/GetCenterAllStudentList', {
        ...baseBody,
        StartNum: startNum,
        DispNum: pageSize
      });
      list = extractStudentList(next);
    } else {
      const next = await apiCall(token, '/api/ATE0010P/GetCenterAllStudentList', {
        ...baseBody,
        Offset: 1 + students.length,
        GetNum: pageSize
      });
      list = extractStudentList(next);
    }
  }

  return students;
}

// Get study result for a student (optional worksheetCD from study.NextWorksheetCD)
async function getStudyResult(token, studentID, classID, classStudentSeq, subjectCD, centerID, worksheetCD) {
  const body = {
    SystemCountryCD: 'USA',
    StudentID: studentID,
    ClassID: classID,
    ClassStudentSeq: classStudentSeq,
    SubjectCD: subjectCD,
    client: clientObject()
  };
  if (centerID) body.CenterID = centerID;
  if (worksheetCD) body.WorksheetCD = worksheetCD;
  
  return apiCall(token, '/api/ATD0010P/GetStudyResultInfoList', body);
}

// Parse LoginID list from text (tab or comma separated, first column = LoginID)
function parseLoginIdList(text) {
  if (!text || typeof text !== 'string') return [];
  const lines = text.replace(/\r/g, '').split('\n');
  const ids = [];
  for (let i = 0; i < lines.length; i++) {
    const trimmed = String(lines[i]).trim();
    if (!trimmed) continue;
    if (i === 0 && /loginid/i.test(trimmed) && /name/i.test(trimmed)) continue;
    let id;
    if (/[\t,]/.test(trimmed)) {
      id = trimmed.split(/[\t,]/)[0].trim();
    } else {
      id = trimmed.split(/\s+/)[0].trim();
    }
    if (!id || /^loginid$/i.test(id)) continue;
    if (!ids.includes(id)) ids.push(id);
  }
  return ids;
}

// Compute lowest planned page from GetStudyResultInfoList response (same logic as Tampermonkey)
function computeLowestFromStudyResult(data) {
  const list = (data && data.StudyUnitInfoList) ? data.StudyUnitInfoList : [];
  const planned = list.filter(u => {
    if (!u) return false;
    if (u.StudyStatus === '6') return false;
    if (u.StudyDate || u.FinishDate) return false;
    return true;
  });
  let minFrom = null, minTo = null, minRow = null;
  for (const u of planned) {
    const from = u.WorksheetNOFrom;
    if (from == null || from === '') continue;
    const fromN = Number(from);
    const toN = (u.WorksheetNOTo != null && u.WorksheetNOTo !== '') ? Number(u.WorksheetNOTo) : null;
    if (isNaN(fromN)) continue;
    if (minFrom === null || fromN < minFrom) {
      minFrom = fromN;
      minTo = (toN != null && !isNaN(toN)) ? toN : null;
      minRow = u.StudyScheduleIndex != null ? u.StudyScheduleIndex : null;
    }
  }
  return { minFrom, minTo, minRow };
}

function subjectName(subjectCD) {
  const cd = String(subjectCD || '');
  if (cd === '010') return 'Math';
  if (cd === '022') return 'Reading';
  return cd ? 'Subject' + cd : cd;
}

// Read stdin to a string (for --list=-)
function readStdin() {
  return new Promise((resolve) => {
    if (process.stdin.isTTY) {
      resolve('');
      return;
    }
    let data = '';
    process.stdin.setEncoding('utf8');
    process.stdin.on('data', chunk => { data += chunk; });
    process.stdin.on('end', () => resolve(data));
  });
}

// Prompt helper
function prompt(question) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  
  return new Promise(resolve => {
    rl.question(question, answer => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

// Main
async function main() {
  const fs = require('fs');
  
  try {
    // Get credentials
    let username = config.username;
    let password = config.password;
    let usePasswordHash = config.passwordHash || config.hash;
    
    if (!username) {
      username = await prompt('Username (LoginID): ');
    }
    if (!password && !usePasswordHash) {
      console.log('Note: Kumon API expects a password hash, not plain password.');
      console.log('You can get the hash from NaviPasswordHash cookie after logging in via browser.');
      const choice = await prompt('Enter (1) password hash OR (2) plain password to try: ');
      if (choice === '1') {
        usePasswordHash = await prompt('Password hash (from NaviPasswordHash cookie): ');
      } else {
        password = await prompt('Plain password (will try, may fail): ');
        process.stdout.write('\x1B[1A\x1B[2K');
      }
    }
    
    const isHash = !!usePasswordHash;
    if (usePasswordHash) {
      password = usePasswordHash;
    }
    
    // Login
    const { token } = await login(username, password, isHash);
    
    // Get instructor info
    console.log('Fetching instructor info...');
    const instructorInfo = await getInstructorInfo(token, username);
    const centerID = instructorInfo.MainCenterID || (instructorInfo.CenterInfoList && instructorInfo.CenterInfoList[0] && instructorInfo.CenterInfoList[0].CenterID);
    const instructorAssistantSec = instructorInfo.InstructorAssistantSec || '2';
    
    console.log('✓ Instructor:', instructorInfo.FullName);
    console.log('✓ CenterID:', centerID);
    
    // ---- Lowest pages (CSV) mode: --lowest --list=file.txt [--subject=010|022|both]
    if (config.lowest !== undefined && config.lowest !== '') {
      const listPath = config.list;
      if (!listPath) {
        console.error('Usage: node kumon-data-puller.js --lowest --list=students.txt [--subject=010|022|both]');
        console.error('  File format: one student per line, LoginID then tab or comma then name (name ignored).');
        process.exit(1);
      }
      let listText;
      if (listPath === '-' || listPath === 'stdin') {
        listText = await readStdin();
      } else {
        listText = fs.readFileSync(listPath, 'utf8');
      }
      const loginIds = parseLoginIdList(listText);
      if (loginIds.length === 0) {
        console.error('No LoginIDs found in list.');
        process.exit(1);
      }
      console.log('Fetching full student list...');
      const allStudents = await getAllStudents(token, centerID, username, instructorAssistantSec);
      console.error(`  API returned ${allStudents.length} students.`);
      const students = allStudents.filter(s => {
        const lid = s.LoginID != null ? String(s.LoginID) : '';
        const sid = s.StudentID != null ? String(s.StudentID) : '';
        return loginIds.includes(lid) || loginIds.includes(sid);
      });
      if (students.length === 0) {
        console.error('  No students from your list matched the API. Check that LoginIDs match.');
        const fromApi = allStudents.slice(0, 5).map(s => s.LoginID || s.StudentID);
        console.error('  Sample IDs from API:', fromApi.join(', '));
        console.error('  Your list (first 5):', loginIds.slice(0, 5).join(', '));
      } else {
        console.error(`  Matched ${students.length} students from your list.`);
      }
      const subjectFilter = config.subject || 'both';
      const wantMath = subjectFilter === 'both' || subjectFilter === '010';
      const wantReading = subjectFilter === 'both' || subjectFilter === '022';
      const rows = [];
      rows.push(['StudentID', 'FullName', 'Subject', 'Level', 'LowestPlannedFrom', 'LowestPlannedTo', 'LowestPlannedIndex'].join('\t'));
      
      let done = 0;
      const total = students.reduce((n, s) => {
        const list = s.StudentStudyInfoList || [];
        return n + list.filter(st => (wantMath && st.SubjectCD === '010') || (wantReading && st.SubjectCD === '022')).length;
      }, 0);
      
      for (const student of students) {
        const list = student.StudentStudyInfoList || [];
        const fullName = student.FullName || student.StudentName || '';
        const studentID = student.StudentID || student.LoginID;
        for (const study of list) {
          if (!study.SubjectCD || study.ClassID == null || study.ClassStudentSeq == null) continue;
          if (study.SubjectCD === '010' && !wantMath) continue;
          if (study.SubjectCD === '022' && !wantReading) continue;
          try {
            done++;
            process.stderr.write(`  [${done}/${total}] ${fullName || studentID} ${subjectName(study.SubjectCD)}...\n`);
            const result = await getStudyResult(
              token,
              studentID,
              study.ClassID,
              study.ClassStudentSeq,
              study.SubjectCD,
              centerID,
              study.NextWorksheetCD
            );
            const lowest = computeLowestFromStudyResult(result);
            rows.push([
              studentID,
              fullName,
              subjectName(study.SubjectCD),
              String(study.NextWorksheetCD || ''),
              lowest.minFrom != null ? lowest.minFrom : '',
              lowest.minTo != null ? lowest.minTo : '',
              lowest.minRow != null ? lowest.minRow : ''
            ].join('\t'));
          } catch (e) {
            process.stderr.write(`  ✗ ${studentID} ${subjectName(study.SubjectCD)}: ${e.message}\n`);
          }
          await new Promise(r => setTimeout(r, 400));
        }
      }
      
      console.log('\n✓ Lowest pages (CSV):');
      console.log(rows.join('\n'));
      if (config.output) {
        fs.writeFileSync(config.output, rows.join('\n'));
        console.error('Saved to ' + config.output);
      }
      return;
    }
    
    // Ask what to pull
    console.log('\nWhat data do you want to pull?');
    console.log('1. Student list (all students)');
    console.log('2. Study results for specific LoginIDs');
    console.log('3. Progress goals for specific LoginIDs');
    console.log('4. All of the above');
    
    const choice = await prompt('Choice (1-4): ');
    
    let students = [];
    if (choice === '1' || choice === '2' || choice === '3' || choice === '4') {
      console.log('\nFetching student list...');
      students = await getAllStudents(token, centerID, username, instructorAssistantSec);
      console.log(`✓ Found ${students.length} students`);
    }
    
    // Export data
    const output = {
      pulledAt: new Date().toISOString(),
      instructor: {
        loginID: username,
        fullName: instructorInfo.FullName,
        centerID: centerID
      },
      students: students.map(s => ({
        LoginID: s.LoginID,
        StudentID: s.StudentID,
        FullName: s.FullName,
        StudentName: s.StudentName,
        ClassID: s.ClassID,
        ClassStudentSeq: s.ClassStudentSeq,
        StudentStudyInfoList: s.StudentStudyInfoList || []
      }))
    };
    
    if (choice === '2' || choice === '4') {
      const loginIdsInput = await prompt('\nEnter LoginIDs (comma-separated): ');
      const loginIds = loginIdsInput.split(',').map(id => id.trim()).filter(Boolean);
      
      output.studyResults = [];
      for (const loginId of loginIds) {
        const student = students.find(s => (s.LoginID || s.StudentID) === loginId);
        if (!student) {
          console.log(`⚠ Student ${loginId} not found`);
          continue;
        }
        
        const studyList = student.StudentStudyInfoList || [];
        for (const study of studyList) {
          try {
            console.log(`  Fetching study result for ${loginId} - Subject ${study.SubjectCD}...`);
            const result = await getStudyResult(
              token,
              student.StudentID || student.LoginID,
              study.ClassID,
              study.ClassStudentSeq,
              study.SubjectCD,
              centerID
            );
            output.studyResults.push({
              loginID: loginId,
              studentID: student.StudentID || student.LoginID,
              subjectCD: study.SubjectCD,
              studyResult: result
            });
            await new Promise(resolve => setTimeout(resolve, 500));
          } catch (e) {
            console.log(`  ✗ Error for ${loginId} Subject ${study.SubjectCD}:`, e.message);
          }
        }
      }
    }
    
    if (choice === '3' || choice === '4') {
      const loginIdsInput = await prompt('\nEnter LoginIDs for progress goals (comma-separated): ');
      const loginIds = loginIdsInput.split(',').map(id => id.trim()).filter(Boolean);
      
      output.progressGoals = [];
      for (const loginId of loginIds) {
        const student = students.find(s => (s.LoginID || s.StudentID) === loginId);
        if (!student) {
          console.log(`⚠ Student ${loginId} not found`);
          continue;
        }
        
        const studyList = student.StudentStudyInfoList || [];
        for (const study of studyList) {
          try {
            console.log(`  Fetching progress goal for ${loginId} - Subject ${study.SubjectCD}...`);
            const result = await apiCall(token, '/api/ATE0020P/GetProgressGoal', {
              SystemCountryCD: 'USA',
              StudentID: student.StudentID || student.LoginID,
              ClassID: study.ClassID,
              ClassStudentSeq: study.ClassStudentSeq,
              SubjectCD: study.SubjectCD,
              client: clientObject()
            });
            output.progressGoals.push({
              loginID: loginId,
              studentID: student.StudentID || student.LoginID,
              subjectCD: study.SubjectCD,
              progressGoal: result
            });
            await new Promise(resolve => setTimeout(resolve, 500));
          } catch (e) {
            console.log(`  ✗ Error for ${loginId} Subject ${study.SubjectCD}:`, e.message);
          }
        }
      }
    }
    
    // Save output
    const outputFile = `kumon-data-${Date.now()}.json`;
    fs.writeFileSync(outputFile, JSON.stringify(output, null, 2));
    console.log(`\n✓ Data saved to ${outputFile}`);
    
  } catch (e) {
    console.error('\n✗ Error:', e.message);
    if (e.stack) console.error(e.stack);
    process.exit(1);
  }
}

main();
