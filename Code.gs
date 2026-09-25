/**

 * シフト表自動生成（固定セル対応・回数設定対応版）

 *

 * ===== 使い方 =====

 * 1. 「シフト」→「ひな形作成/更新」で、シフト表に従業員×日付の表を作る

 * 2. シフト表に、先に決めたい人・日付の業務を手入力する

 * 3. そのセルを選択して「シフト」→「選択セルを固定」を押す

 * 4. 「シフト」→「回数設定ひな形作成/更新」で、回数設定シートを作る

 * 5. 「回数設定」シートで、出勤日数ごとの区分1/区分2の目標回数を設定する

 * 6. 「シフト」→「生成」を押す

 *

 * ===== シート構成 =====

 * ① 従業員（SHEET_EMP）:

 *    A=名前, B=区分(午前/午後), C=空白禁止（任意：TRUE/はい/1/yes/◯ など）

 *

 * ② 業務（SHEET_JOB）:

 *    A列に業務名（A2以降）

 *

 * ③ 設定（SHEET_CFG）:

 *    B1=開始日, B2=終了日, B3=間隔(日), B5以降=会社休業日（任意）

 *

 * ④ フォーム回答（SHEET_FORM）:

 *    「氏名」「休暇希望日」列があること（同じ氏名は最新回答のみ採用）

 *

 * ⑤ 業務分類（SHEET_CAT）:

 *    A=業務, B=分類（区分1/区分2）

 *    ※ここは従来通り、同じ分類の連続を避ける補助に使う

 *

 * ⑥ 回数設定（SHEET_QUOTA）:

 *    A=出勤日数, B=区分1, C=区分2

 *    ※業務の分類（区分1/区分2）は、既存の「業務分類」シートをそのまま使用する

 *

 * ⑦ シフト表（SHEET_OUT）:

 *    出力先。ここに先に入力した固定セルも置く

 *

 * ⑧ 警告（SHEET_WARN）:

 *    警告出力先

 */



const SHEET_EMP  = "従業員";

const SHEET_JOB  = "業務";

const SHEET_CFG  = "設定";

const SHEET_OUT  = "シフト表";

const SHEET_FORM = "フォーム回答同期（初回許可）";

const SHEET_WARN = "警告";

const SHEET_CAT  = "業務分類";

const SHEET_QUOTA = "回数設定";



const TZ = "Asia/Tokyo";



const FIXED_NOTE = "固定";

const FIXED_BG = "#fff2cc";

const AUTO_BG = "#ffffff";



const IMPLICIT_FIXED_MAX_RATIO = 0.40;



const JOB_CONFLICT_GROUPS = [ // Optional job groups for spacing rules; empty by default.

  [],

  [],

  [],

  [],

];

const JOB_CONFLICT_MIN_GAP_DAYS = 3;



const COUNT_CAT_PRIMARY = "区分1";

const COUNT_CAT_SECONDARY = "区分2";

const COUNT_CATEGORY_NAMES = [COUNT_CAT_PRIMARY, COUNT_CAT_SECONDARY];



function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu("シフト")

    .addItem("ひな形作成/更新", "prepareShiftTemplate")

    .addItem("回数設定ひな形作成/更新", "prepareQuotaTemplate")

    .addSeparator()

    .addItem("選択セルを固定", "markSelectedCellsAsFixed")

    .addItem("選択セルの固定を解除", "unmarkSelectedCellsAsFixed")

    .addSeparator()

    .addItem("生成", "generateShift")

    .addSeparator()

    .addItem("Webアプリ接続設定", "setupWebAppConnection")

    .addToUi();

}



function prepareShiftTemplate(ssOverride, silent) {

  const ss = ssOverride || SpreadsheetApp.getActiveSpreadsheet();

  return withShiftLock_(function() {
    return prepareShiftTemplateCore_(ss, silent);
  });
}

function prepareShiftTemplateCore_(ss, silent) {

  const employees = readEmployees_(ss);

  const cfg = readConfig_(ss);

  const dates = enumerateDates_(cfg.startDate, cfg.endDate);



  const fixedRead = readFixedEntriesFromOutput_(ss, employees, dates);

  const fixedByCell = new Map();

  for (const f of fixedRead.entries) {

    fixedByCell.set(cellKey_(f.ei, f.di), f.value);

  }



  const out = ss.getSheetByName(SHEET_OUT) || ss.insertSheet(SHEET_OUT);

  writeShiftOutput_(out, employees, dates, createEmptyGrid_(employees.length, dates.length), fixedByCell);



  if (!silent) SpreadsheetApp.getUi().alert("シフト表のひな形を作成/更新しました。先に決めたいセルを入力して、必要に応じて『選択セルを固定』を押してください。");

  return { employeeCount: employees.length, dateCount: dates.length };

}



function prepareQuotaTemplate(ssOverride, silent) {

  const ss = ssOverride || SpreadsheetApp.getActiveSpreadsheet();

  return withShiftLock_(function() {
    return prepareQuotaTemplateCore_(ss, silent);
  });
}

function prepareQuotaTemplateCore_(ss, silent) {

  const sh = ss.getSheetByName(SHEET_QUOTA) || ss.insertSheet(SHEET_QUOTA);



  sh.getRange(1, 1, 1, 3).setValues([["出勤日数", "区分1", "区分2"]]);



  const maxRows = Math.max(1, sh.getMaxRows() - 1);

  const existingQuota = sh.getRange(2, 1, maxRows, 3).getValues()

    .some(row => row.some(v => String(v || "").trim() !== ""));



  if (!existingQuota) {

    sh.getRange(2, 1, 2, 3).setValues([

      [10, 5, 5],

      [12, 6, 6],

    ]);

  }



  sh.setFrozenRows(1);

  sh.autoResizeColumns(1, 3);



  if (!silent) SpreadsheetApp.getUi().alert(

    "回数設定シートを作成/更新しました。\n\n" +

    "A:C に出勤日数ごとの区分1/区分2の目標回数を設定してください。\n" +

    "業務ごとの区分1/区分2分類は、既存の「業務分類」シートを使用します。"

  );

  return { rowCount: Math.max(0, sh.getLastRow() - 1) };

}



function markSelectedCellsAsFixed() {

  return withShiftLock_(function() {
    return markSelectedCellsAsFixedCore_();
  });
}

function markSelectedCellsAsFixedCore_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sh = ss.getActiveSheet();

  if (!sh || sh.getName() !== SHEET_OUT) {

    SpreadsheetApp.getUi().alert(`「${SHEET_OUT}」シートで固定したいセルを選択してください。`);

    return;

  }



  const ranges = getSelectedRanges_(sh);

  if (ranges.length === 0) return;



  let count = 0;



  for (const range of ranges) {

    const values = range.getValues();

    const notes = range.getNotes();

    const backgrounds = range.getBackgrounds();



    for (let r = 0; r < values.length; r++) {

      for (let c = 0; c < values[0].length; c++) {

        const row = range.getRow() + r;

        const col = range.getColumn() + c;

        if (row < 2 || col < 2) continue;



        const value = normalizeCellValue_(values[r][c]);

        if (!value) continue;



        notes[r][c] = FIXED_NOTE;

        backgrounds[r][c] = FIXED_BG;

        count++;

      }

    }



    range.setNotes(notes);

    range.setBackgrounds(backgrounds);

  }



  SpreadsheetApp.getUi().alert(`${count}個のセルを固定しました。`);

}



function unmarkSelectedCellsAsFixed() {

  return withShiftLock_(function() {
    return unmarkSelectedCellsAsFixedCore_();
  });
}

function unmarkSelectedCellsAsFixedCore_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sh = ss.getActiveSheet();

  if (!sh || sh.getName() !== SHEET_OUT) {

    SpreadsheetApp.getUi().alert(`「${SHEET_OUT}」シートで固定解除したいセルを選択してください。`);

    return;

  }



  const ranges = getSelectedRanges_(sh);

  if (ranges.length === 0) return;



  let count = 0;



  for (const range of ranges) {

    const notes = range.getNotes();

    const backgrounds = range.getBackgrounds();



    for (let r = 0; r < notes.length; r++) {

      for (let c = 0; c < notes[0].length; c++) {

        const row = range.getRow() + r;

        const col = range.getColumn() + c;

        if (row < 2 || col < 2) continue;



        if (isFixedMark_(notes[r][c], backgrounds[r][c])) {

          notes[r][c] = "";

          backgrounds[r][c] = AUTO_BG;

          count++;

        }

      }

    }



    range.setNotes(notes);

    range.setBackgrounds(backgrounds);

  }



  SpreadsheetApp.getUi().alert(`${count}個のセルの固定を解除しました。`);

}



function getSelectedRanges_(sh) {

  const rangeList = sh.getActiveRangeList();

  if (rangeList) return rangeList.getRanges();



  const range = sh.getActiveRange();

  return range ? [range] : [];

}



function generateShift(ssOverride) {

  const ss = ssOverride || SpreadsheetApp.getActiveSpreadsheet();

  return withShiftLock_(function() {
    return generateShiftCore_(ss);
  });
}

function generateShiftCore_(ss) {



  const employees = readEmployees_(ss);

  const employeeNames = employees.map(e => e.name);

  const nameToIndex = new Map(employeeNames.map((n, i) => [n, i]));



  const jobs = readSingleColumn_(ss, SHEET_JOB, 1, 2);

  if (employees.length === 0) throw new Error("従業員シートに従業員がいません（A列に名前）。");

  if (jobs.length === 0) throw new Error("業務シートに業務がありません（A2以降）。");



  const jobsSet = new Set(jobs);



  const cfg = readConfig_(ss);

  const start = cfg.startDate;

  const end = cfg.endDate;

  const intervalDays = cfg.intervalDays;



  const dates = enumerateDates_(start, end);

  const companyHolidays = readCompanyHolidays_(ss, cfg, start);

  const vacationsByName = readVacationsFromFormLatest_(ss, start, end);



  const jobCategory = readJobCategoryMap_(ss); // Map(job -> "区分1"|"区分2")



  // 回数制御でも、既存の「業務分類」シートの区分1/区分2をそのまま使う

  const countCategory = jobCategory; // Map(job -> "区分1"|"区分2")



  const quotaRules = readCategoryQuotaRules_(ss); // Map(workDays -> Map("区分1"/"区分2" -> target))



  if (quotaRules.size > 0) {

    for (const job of jobs) {

      if (!jobCategory.has(job)) {

        throw new Error(

          `業務分類シートに、業務「${job}」の分類がありません。` +

          `「区分1」または「区分2」を設定してください。回数設定はこの分類を使います。`

        );

      }

    }

  }



  const conflictGroupOfJob = new Map();

  JOB_CONFLICT_GROUPS.forEach((arr, gi) => arr.forEach(j => conflictGroupOfJob.set(j, gi)));



  const fixedRead = readFixedEntriesFromOutput_(ss, employees, dates);



  const warn = [];

  for (const w of fixedRead.warnRows) warn.push(w);



  const rows = employees.length;

  const cols = dates.length;

  const grid = createEmptyGrid_(rows, cols);



  const fixedByCell = new Map();

  const fixedByDay = new Map();



  for (const f of fixedRead.entries) {

    const key = cellKey_(f.ei, f.di);

    if (fixedByCell.has(key)) {

      warn.push([formatMD_(dates[f.di]), `固定セルが重複しています: ${f.name}`]);

      continue;

    }



    grid[f.ei][f.di] = f.value;

    fixedByCell.set(key, f.value);



    if (!fixedByDay.has(f.di)) fixedByDay.set(f.di, []);

    fixedByDay.get(f.di).push({ ei: f.ei, name: f.name, value: f.value });

  }



  for (let di = 0; di < dates.length; di++) {

    const key = dateKey_(dates[di]);

    const isCompanyHoliday = companyHolidays.has(key);



    for (let ei = 0; ei < employees.length; ei++) {

      const name = employees[ei].name;

      const vacSet = vacationsByName.get(name);

      const isVacation = vacSet ? vacSet.has(key) : false;

      const isAutoRest = isCompanyHoliday || isVacation;

      const fixed = fixedByCell.has(cellKey_(ei, di));

      const current = grid[ei][di];



      if (fixed && isAutoRest && current !== "休") {

        const reason = isCompanyHoliday ? "会社休業日" : "休暇希望日";

        warn.push([formatMD_(dates[di]), `固定セルが${reason}と重複しています: ${name}=${current}`]);

        continue;

      }



      if (!fixed && isAutoRest) grid[ei][di] = "休";

    }

  }



  const categoryTargetsByName = buildCategoryTargets_(employees, dates, grid, quotaRules, warn);



  validateFixedSameDayJobDuplication_(warn, fixedByDay, dates, jobsSet);



  const lastDone = new Map();

  const totalAssigned = new Map();

  const jobCount = new Map();

  const categoryCount = new Map();

  const lastCategory = new Map();

  const lastConflictGroupDone = new Map();



  employeeNames.forEach(n => {

    lastDone.set(n, new Map());

    totalAssigned.set(n, 0);

    jobCount.set(n, new Map());

    categoryCount.set(n, new Map());

    lastCategory.set(n, null);

    lastConflictGroupDone.set(n, new Map());

  });



  const morningIdxAll = [];

  const afternoonIdxAll = [];

  for (let ei = 0; ei < employees.length; ei++) {

    if (employees[ei].group === "午前") morningIdxAll.push(ei);

    else afternoonIdxAll.push(ei);

  }



  for (let di = 0; di < dates.length; di++) {

    const day = dates[di];

    const dayLabel = formatMD_(day);

    const key = dateKey_(day);



    const morningIdx = morningIdxAll;

    const afternoonIdx = afternoonIdxAll;



    const assignedToday = new Set();

    const usedJobsToday = new Set();



    const todaysFixed = fixedByDay.get(di) || [];

    for (const f of todaysFixed) {

      const value = f.value;

      if (value === "休") continue;



      assignedToday.add(f.name);



      if (!jobsSet.has(value)) {

        warn.push([dayLabel, `固定セルの値が業務一覧にありません: ${f.name}=${value}`]);

        continue;

      }



      if (usedJobsToday.has(value)) {

        warn.push([dayLabel, `同じ業務が固定セルで複数人に入っています: ${value}`]);

      }

      usedJobsToday.add(value);



      validateFixedAgainstHistory_(

        warn,

        f.name,

        value,

        di,

        dayLabel,

        intervalDays,

        lastDone,

        lastConflictGroupDone,

        conflictGroupOfJob

      );



      registerAssignment_(

        f.name,

        value,

        di,

        jobCategory,

        countCategory,

        conflictGroupOfJob,

        lastDone,

        totalAssigned,

        jobCount,

        categoryCount,

        lastCategory,

        lastConflictGroupDone

      );

    }



    if (companyHolidays.has(key)) continue;



    let mRest = 0, aRest = 0;

    for (const ei of morningIdx) if (grid[ei][di] === "休") mRest++;

    for (const ei of afternoonIdx) if (grid[ei][di] === "休") aRest++;



    let mAssignedCount = countNonRestNonBlankInGroup_(grid, di, morningIdx);

    let aAssignedCount = countNonRestNonBlankInGroup_(grid, di, afternoonIdx);



    if (2 - mRest < 0) warn.push([dayLabel, `午前組：休が${mRest}人で条件(休+空白<=2)が不可能`]);

    if (2 - aRest < 0) warn.push([dayLabel, `午後組：休が${aRest}人で条件(休+空白<=2)が不可能`]);



    const mNeedAssign = Math.max(0, Math.min(morningIdx.length - 2, morningIdx.length - mRest));

    const aNeedAssign = Math.max(0, Math.min(afternoonIdx.length - 2, afternoonIdx.length - aRest));



    const mAvailable = morningIdx

      .filter(ei => grid[ei][di] === "")

      .map(ei => employees[ei].name);

    const aAvailable = afternoonIdx

      .filter(ei => grid[ei][di] === "")

      .map(ei => employees[ei].name);



    const mustSet = new Set(

      employees

        .filter((e, ei) => e.must && grid[ei][di] === "")

        .map(e => e.name)

    );



    const remainingJobsInitial = jobs.filter(j => !usedJobsToday.has(j));

    if (mustSet.size > remainingJobsInitial.length) {

      warn.push([dayLabel, `空白禁止の未割当者が${mustSet.size}人ですが、残り業務が${remainingJobsInitial.length}個しかなく、全員に割当は不可能です。`]);

    }



    const remainingJobs = [...remainingJobsInitial];



    while (remainingJobs.length > 0) {

      let bestIdx = -1;

      let bestScore = Infinity;



      for (let j = 0; j < remainingJobs.length; j++) {

        const job = remainingJobs[j];

        const countCat = countCategory.get(job) || null;



        const candM_all = eligibleCandidates_(job, mAvailable, assignedToday, lastDone, lastConflictGroupDone, conflictGroupOfJob, di, intervalDays);

        const candA_all = eligibleCandidates_(job, aAvailable, assignedToday, lastDone, lastConflictGroupDone, conflictGroupOfJob, di, intervalDays);



        const all = candM_all.concat(candA_all);

        if (all.length === 0) continue;



        const unexpCount = countUnexperienced_(all, job, jobCount);

        const quotaNeedCount = countPositiveCategoryDeficit_(all, countCat, categoryTargetsByName, categoryCount);



        let score;



        if (mustSet.size > 0) {

          const mustCandidates = all.filter(n => mustSet.has(n));

          if (mustCandidates.length > 0) {

            const mustQuotaNeedCount = countPositiveCategoryDeficit_(mustCandidates, countCat, categoryTargetsByName, categoryCount);

            const mustUnexpCount = countUnexperienced_(mustCandidates, job, jobCount);



            if (mustQuotaNeedCount > 0) score = 0 * 1000000000 + mustQuotaNeedCount * 1000 + all.length;

            else if (mustUnexpCount > 0) score = 1 * 1000000000 + mustUnexpCount * 1000 + all.length;

            else score = 2 * 1000000000 + all.length;

          } else {

            score = 9 * 1000000000 + all.length;

          }

        } else {

          if (quotaNeedCount > 0) score = 0 * 1000000000 + quotaNeedCount * 1000 + all.length;

          else if (unexpCount > 0) score = 1 * 1000000000 + unexpCount * 1000 + all.length;

          else score = 2 * 1000000000 + all.length;

        }



        if (score < bestScore) {

          bestScore = score;

          bestIdx = j;

        }

      }



      if (bestIdx === -1) break;



      const job = remainingJobs.splice(bestIdx, 1)[0];

      const jobCat = jobCategory.get(job) || null;

      const countCat = countCategory.get(job) || null;



      const candM = eligibleCandidates_(job, mAvailable, assignedToday, lastDone, lastConflictGroupDone, conflictGroupOfJob, di, intervalDays);

      const candA = eligibleCandidates_(job, aAvailable, assignedToday, lastDone, lastConflictGroupDone, conflictGroupOfJob, di, intervalDays);



      const allCandidates = candM.concat(candA);

      if (allCandidates.length === 0) continue;



      let baseCandidates = allCandidates;



      if (mustSet.size > 0) {

        const mustCandidates = allCandidates.filter(n => mustSet.has(n));

        if (mustCandidates.length > 0) baseCandidates = mustCandidates;

      }



      const baseSet = new Set(baseCandidates);

      const candMBase = candM.filter(n => baseSet.has(n));

      const candABase = candA.filter(n => baseSet.has(n));



      let pool = [];



      const quotaCandidates = baseCandidates.filter(n =>

        categoryDeficit_(n, countCat, categoryTargetsByName, categoryCount) > 0

      );



      if (quotaCandidates.length > 0) {

        pool = quotaCandidates;

      } else {

        const unexpCandidates = baseCandidates.filter(n => (jobCount.get(n).get(job) || 0) === 0);



        if (unexpCandidates.length > 0) {

          pool = unexpCandidates;

        } else {

          if (mAssignedCount < mNeedAssign && candMBase.length > 0) pool = candMBase;

          else if (aAssignedCount < aNeedAssign && candABase.length > 0) pool = candABase;

          else pool = baseCandidates;

        }

      }



      if (pool.length === 0) continue;



      pool.sort((a, b) => {

        const aQuota = categoryQuotaSortScore_(a, countCat, categoryTargetsByName, categoryCount);

        const bQuota = categoryQuotaSortScore_(b, countCat, categoryTargetsByName, categoryCount);

        if (aQuota !== bQuota) return aQuota - bQuota;



        const aJC = jobCount.get(a).get(job) || 0;

        const bJC = jobCount.get(b).get(job) || 0;

        if (aJC !== bJC) return aJC - bJC;



        const aNeed = groupNeedPenalty_(a, employees, nameToIndex, mAssignedCount, mNeedAssign, aAssignedCount, aNeedAssign);

        const bNeed = groupNeedPenalty_(b, employees, nameToIndex, mAssignedCount, mNeedAssign, aAssignedCount, aNeedAssign);

        if (aNeed !== bNeed) return aNeed - bNeed;



        if (jobCat) {

          const aLast = lastCategory.get(a);

          const bLast = lastCategory.get(b);

          const aPenalty = (aLast && aLast === jobCat) ? 1 : 0;

          const bPenalty = (bLast && bLast === jobCat) ? 1 : 0;

          if (aPenalty !== bPenalty) return aPenalty - bPenalty;

        }



        const aTA = totalAssigned.get(a);

        const bTA = totalAssigned.get(b);

        if (aTA !== bTA) return aTA - bTA;



        return a.localeCompare(b, "ja");

      });



      const chosen = pool[0];

      const ei = nameToIndex.get(chosen);



      grid[ei][di] = job;

      assignedToday.add(chosen);

      usedJobsToday.add(job);



      registerAssignment_(

        chosen,

        job,

        di,

        jobCategory,

        countCategory,

        conflictGroupOfJob,

        lastDone,

        totalAssigned,

        jobCount,

        categoryCount,

        lastCategory,

        lastConflictGroupDone

      );



      if (employees[ei].group === "午前") mAssignedCount++;

      else aAssignedCount++;



      mustSet.delete(chosen);

    }



    const mBlank = countBlankInGroup_(grid, di, morningIdx);

    const aBlank = countBlankInGroup_(grid, di, afternoonIdx);

    if (mRest + mBlank > 2) warn.push([dayLabel, `午前組：休+空白=${mRest + mBlank}（>2）`]);

    if (aRest + aBlank > 2) warn.push([dayLabel, `午後組：休+空白=${aRest + aBlank}（>2）`]);



    const mustBlank = employees

      .map((e, ei) => ({ e, ei }))

      .filter(x => x.e.must && grid[x.ei][di] !== "休" && grid[x.ei][di] === "")

      .map(x => x.e.name);



    if (mustBlank.length > 0) {

      warn.push([dayLabel, `空白禁止なのに未割当（空白）が残りました: ${mustBlank.join(", ")}`]);

    }

  }



  appendUnexperiencedWarnings_(warn, employees, jobs, grid, dates, jobCount, conflictGroupOfJob);

  appendCategoryQuotaWarnings_(warn, employees, categoryTargetsByName, categoryCount);



  const out = ss.getSheetByName(SHEET_OUT) || ss.insertSheet(SHEET_OUT);

  writeShiftOutput_(out, employees, dates, grid, fixedByCell);

  writeWarnings_(ss, warn);

  return {
    warningCount: warn.length,
    dateCount: dates.length,
    employeeCount: employees.length,
  };

}



function writeShiftOutput_(out, employees, dates, grid, fixedByCell) {

  out.clear();



  const employeeNames = employees.map(e => e.name);

  const rows = employees.length;

  const cols = dates.length;



  const header = ["従業員"].concat(dates.map(d => formatMD_(d)));

  out.getRange(1, 1, 1, header.length).setValues([header]);



  if (rows > 0) {

    out.getRange(2, 1, rows, 1).setValues(employeeNames.map(n => [n]));

  }



  if (rows > 0 && cols > 0) {

    out.getRange(2, 2, rows, cols).setValues(grid);



    const backgrounds = [];

    const notes = [];

    const weights = [];



    for (let ei = 0; ei < rows; ei++) {

      const bgRow = [];

      const noteRow = [];

      const weightRow = [];

      for (let di = 0; di < cols; di++) {

        if (fixedByCell.has(cellKey_(ei, di))) {

          bgRow.push(FIXED_BG);

          noteRow.push(FIXED_NOTE);

          weightRow.push("bold");

        } else {

          bgRow.push(AUTO_BG);

          noteRow.push("");

          weightRow.push("normal");

        }

      }

      backgrounds.push(bgRow);

      notes.push(noteRow);

      weights.push(weightRow);

    }



    out.getRange(2, 2, rows, cols)

      .setBackgrounds(backgrounds)

      .setNotes(notes)

      .setFontWeights(weights);

  }



  out.setFrozenRows(1);

  out.setFrozenColumns(1);

  out.autoResizeColumns(1, cols + 1);

}



function readFixedEntriesFromOutput_(ss, employees, dates) {

  const sh = ss.getSheetByName(SHEET_OUT);

  const result = { entries: [], warnRows: [] };



  if (!sh) return result;

  if (sh.getLastRow() < 2 || sh.getLastColumn() < 2) return result;



  const lastRow = sh.getLastRow();

  const lastCol = sh.getLastColumn();

  const dataRows = lastRow - 1;

  const dataCols = lastCol - 1;



  const employeeIndex = new Map(employees.map((e, i) => [e.name, i]));

  const dateIndexByKey = new Map(dates.map((d, i) => [dateKey_(d), i]));

  const startRef = dates[0];



  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  const names = sh.getRange(1, 1, lastRow, 1).getValues().map(r => normalizeCellValue_(r[0]));



  const colToDi = new Map();

  for (let col = 2; col <= lastCol; col++) {

    const dt = normalizeToDate_(header[col - 1], startRef);

    if (!dt) continue;

    const k = dateKey_(dt);

    if (dateIndexByKey.has(k)) colToDi.set(col, dateIndexByKey.get(k));

  }



  const rowToEi = new Map();

  for (let row = 2; row <= lastRow; row++) {

    const name = names[row - 1];

    if (!name) continue;

    if (employeeIndex.has(name)) rowToEi.set(row, employeeIndex.get(name));

  }



  if (colToDi.size === 0 || rowToEi.size === 0) return result;



  const values = sh.getRange(2, 2, dataRows, dataCols).getValues();

  const notes = sh.getRange(2, 2, dataRows, dataCols).getNotes();

  const backgrounds = sh.getRange(2, 2, dataRows, dataCols).getBackgrounds();



  const allNonEmpty = [];

  const marked = [];



  for (let r = 0; r < dataRows; r++) {

    const row = r + 2;

    if (!rowToEi.has(row)) continue;

    const ei = rowToEi.get(row);

    const name = employees[ei].name;



    for (let c = 0; c < dataCols; c++) {

      const col = c + 2;

      if (!colToDi.has(col)) continue;

      const di = colToDi.get(col);



      const value = normalizeCellValue_(values[r][c]);

      if (!value) continue;



      const entry = { ei, di, name, value, row, col };

      allNonEmpty.push(entry);



      if (isFixedMark_(notes[r][c], backgrounds[r][c])) marked.push(entry);

    }

  }



  if (marked.length > 0) {

    result.entries = marked;

    return result;

  }



  if (allNonEmpty.length === 0) return result;



  const totalCells = employees.length * dates.length;

  const implicitLimit = Math.max(1, Math.floor(totalCells * IMPLICIT_FIXED_MAX_RATIO));



  if (allNonEmpty.length <= implicitLimit) {

    result.entries = allNonEmpty;

    result.warnRows.push(["-", `固定マークが無かったため、既存の非空セル${allNonEmpty.length}個を固定セルとして扱いました。次回以降は黄色の固定セルのみ固定扱いになります。`]);

  } else {

    result.warnRows.push(["-", `シフト表に既存入力が${allNonEmpty.length}個ありますが、固定マークが無く数が多いため、旧出力と判断して固定扱いしませんでした。固定したいセルは選択して「選択セルを固定」を押してください。`]);

  }



  return result;

}



function validateFixedSameDayJobDuplication_(warn, fixedByDay, dates, jobsSet) {

  for (const [di, arr] of fixedByDay.entries()) {

    const map = new Map();

    for (const f of arr) {

      if (f.value === "休") continue;

      if (!jobsSet.has(f.value)) continue;

      if (!map.has(f.value)) map.set(f.value, []);

      map.get(f.value).push(f.name);

    }



    for (const [job, names] of map.entries()) {

      if (names.length > 1) {

        warn.push([formatMD_(dates[di]), `固定セルで同じ業務「${job}」が複数人に指定されています: ${names.join(", ")}`]);

      }

    }

  }

}



function validateFixedAgainstHistory_(

  warn,

  name,

  job,

  dayIndex,

  dayLabel,

  intervalDays,

  lastDone,

  lastConflictGroupDone,

  conflictGroupOfJob

) {

  const last = lastDone.get(name).get(job);

  if (typeof last === "number") {

    const diff = dayIndex - last;

    if (diff <= intervalDays) {

      warn.push([dayLabel, `固定セルが同一業務の間隔制約に違反しています: ${name}=${job}（前回から${diff}日）`]);

    }

  }



  const g = conflictGroupOfJob.get(job);

  if (typeof g === "number") {

    const lastG = lastConflictGroupDone.get(name).get(g);

    if (typeof lastG === "number") {

      const diffG = dayIndex - lastG;

      if (diffG <= JOB_CONFLICT_MIN_GAP_DAYS) {

        warn.push([dayLabel, `固定セルが業務グループ間隔に違反しています: ${name}=${job}（前回同グループから${diffG}日）`]);

      }

    }

  }

}



function isFixedMark_(note, background) {

  const n = String(note || "").trim();

  const bg = String(background || "").toLowerCase();

  return n.indexOf(FIXED_NOTE) !== -1 || bg === FIXED_BG.toLowerCase();

}



function cellKey_(ei, di) {

  return `${ei}|${di}`;

}



function normalizeCellValue_(v) {

  if (v == null) return "";

  if (v instanceof Date && !isNaN(v.getTime())) return formatMD_(v);

  return String(v).trim();

}



function registerAssignment_(

  name,

  job,

  dayIndex,

  jobCategory,

  countCategory,

  conflictGroupOfJob,

  lastDone,

  totalAssigned,

  jobCount,

  categoryCount,

  lastCategory,

  lastConflictGroupDone

) {

  totalAssigned.set(name, totalAssigned.get(name) + 1);

  lastDone.get(name).set(job, dayIndex);



  const jcMap = jobCount.get(name);

  jcMap.set(job, (jcMap.get(job) || 0) + 1);



  const countCat = countCategory.get(job) || null;

  if (countCat) {

    const ccMap = categoryCount.get(name);

    ccMap.set(countCat, (ccMap.get(countCat) || 0) + 1);

  }



  const jobCat = jobCategory.get(job) || null;

  if (jobCat) lastCategory.set(name, jobCat);



  const jobG = conflictGroupOfJob.get(job);

  if (typeof jobG === "number") lastConflictGroupDone.get(name).set(jobG, dayIndex);

}



function readConfig_(ss) {

  const sh = ss.getSheetByName(SHEET_CFG);

  if (!sh) throw new Error("設定シートが見つかりません。");



  const start = sh.getRange("B1").getValue();

  const end = sh.getRange("B2").getValue();

  const interval = sh.getRange("B3").getValue();



  if (!(start instanceof Date) || isNaN(start.getTime())) throw new Error("設定!B1 の開始日が日付ではありません。");

  if (!(end instanceof Date) || isNaN(end.getTime())) throw new Error("設定!B2 の終了日が日付ではありません。");



  const intervalDays = Number(interval);

  if (!Number.isFinite(intervalDays) || intervalDays < 0) throw new Error("設定!B3 の間隔(日)が不正です。");



  return {

    startDate: start,

    endDate: end,

    intervalDays: Math.floor(intervalDays),

    holidayStartRow: 5,

  };

}



function readCompanyHolidays_(ss, cfg, yearRefDate) {

  const sh = ss.getSheetByName(SHEET_CFG);

  const startRow = cfg.holidayStartRow;

  const lastRow = sh.getLastRow();

  const set = new Set();



  if (lastRow < startRow) return set;



  const values = sh.getRange(startRow, 2, lastRow - startRow + 1, 1).getValues();

  for (const [v] of values) {

    const dt = normalizeToDate_(v, yearRefDate);

    if (!dt) continue;

    set.add(dateKey_(dt));

  }

  return set;

}



function readVacationsFromFormLatest_(ss, start, end) {

  const sh = ss.getSheetByName(SHEET_FORM);

  const map = new Map();

  if (!sh) return map;



  const data = sh.getDataRange().getValues();

  if (data.length <= 1) return map;



  const header = data[0].map(String);

  const idxName = header.indexOf("氏名");

  const idxVac = header.indexOf("休暇希望日");



  if (idxName === -1 || idxVac === -1) {

    throw new Error(`フォーム回答シートに「氏名」「休暇希望日」列が見つかりません。現在のヘッダ: ${header.join(", ")}`);

  }



  const startKey = dateKey_(start);

  const endKey = dateKey_(end);



  for (let i = data.length - 1; i >= 1; i--) {

    const row = data[i];

    const name = String(row[idxName] || "").trim();

    if (!name) continue;

    if (map.has(name)) continue;



    const set = new Set();

    const vacRaw = row[idxVac];



    if (vacRaw instanceof Date) {

      const k = dateKey_(vacRaw);

      if (k >= startKey && k <= endKey) set.add(k);

    } else {

      const s = String(vacRaw || "").trim();

      if (s) {

        const parts = s.split(/,|\n/).map(x => x.trim()).filter(Boolean);

        for (const p of parts) {

          const dt = parseMD_(p, start);

          if (!dt) continue;

          const k = dateKey_(dt);

          if (k >= startKey && k <= endKey) set.add(k);

        }

      }

    }



    map.set(name, set);

  }



  return map;

}



function readEmployees_(ss) {

  const sh = ss.getSheetByName(SHEET_EMP);

  if (!sh) throw new Error("従業員シートが見つかりません。");



  const last = sh.getLastRow();

  if (last < 2) throw new Error("従業員シートにデータがありません。");



  const values = sh.getRange(2, 1, last - 1, 3).getValues();

  const employees = [];



  for (const [nameRaw, groupRaw, mustRaw] of values) {

    const name = String(nameRaw || "").trim();

    const group = String(groupRaw || "").trim();

    if (!name) continue;



    if (group !== "午前" && group !== "午後") {

      throw new Error(`従業員「${name}」の区分が「午前/午後」ではありません: ${group}`);

    }



    const must = parseBool_(mustRaw);

    employees.push({ name, group, must });

  }

  return employees;

}



function parseBool_(v) {

  if (v === true) return true;

  if (v === false || v == null) return false;

  const s = String(v).trim().toLowerCase();

  return (

    s === "true" || s === "1" || s === "yes" || s === "y" ||

    s === "はい" || s === "〇" || s === "◯"

  );

}



function readSingleColumn_(ss, sheetName, col, startRow) {

  const sh = ss.getSheetByName(sheetName);

  if (!sh) throw new Error(`${sheetName} シートが見つかりません。`);



  const last = sh.getLastRow();

  if (last < startRow) return [];



  return sh.getRange(startRow, col, last - startRow + 1, 1)

    .getValues()

    .map(r => String(r[0] || "").trim())

    .filter(Boolean);

}



function readJobCategoryMap_(ss) {

  const sh = ss.getSheetByName(SHEET_CAT);

  const map = new Map();



  if (!sh) return map;



  const data = sh.getDataRange().getValues();

  if (data.length <= 1) return map;



  for (let i = 1; i < data.length; i++) {

    const job = String(data[i][0] || "").trim();

    const cat = String(data[i][1] || "").trim();

    if (!job) continue;

    if (cat !== "区分1" && cat !== "区分2") continue;

    map.set(job, cat);

  }

  return map;

}



function readCategoryQuotaRules_(ss) {

  const sh = ss.getSheetByName(SHEET_QUOTA);

  const map = new Map();



  // シートがなければ回数制御なし

  if (!sh) return map;



  const lastRow = sh.getLastRow();

  if (lastRow < 2) return map;



  const values = sh.getRange(2, 1, lastRow - 1, 3).getValues(); // A:C



  for (let i = 0; i < values.length; i++) {

    const workDaysRaw = values[i][0];

    const primaryRaw = values[i][1];

    const secondaryRaw = values[i][2];



    if (workDaysRaw === "" || workDaysRaw == null) continue;



    const workDays = Number(workDaysRaw);

    const primaryCount = Number(primaryRaw);

    const secondaryCount = Number(secondaryRaw);



    if (!Number.isInteger(workDays) || workDays < 0) {

      throw new Error(`回数設定シート ${i + 2}行目：出勤日数が不正です。`);

    }

    if (!Number.isInteger(primaryCount) || primaryCount < 0) {

      throw new Error(`回数設定シート ${i + 2}行目：区分1の回数が不正です。`);

    }

    if (!Number.isInteger(secondaryCount) || secondaryCount < 0) {

      throw new Error(`回数設定シート ${i + 2}行目：区分2の回数が不正です。`);

    }

    if (primaryCount + secondaryCount !== workDays) {

      throw new Error(

        `回数設定シート ${i + 2}行目：区分1と区分2の回数の合計が出勤日数と一致していません。` +

        ` 出勤日数=${workDays}, 区分1=${primaryCount}, 区分2=${secondaryCount}`

      );

    }



    const quota = new Map();

    quota.set(COUNT_CAT_PRIMARY, primaryCount);

    quota.set(COUNT_CAT_SECONDARY, secondaryCount);

    map.set(workDays, quota);

  }



  return map;

}



function buildCategoryTargets_(employees, dates, grid, quotaRules, warn) {

  const targetByName = new Map();



  if (!quotaRules || quotaRules.size === 0) {

    for (const e of employees) targetByName.set(e.name, null);

    return targetByName;

  }



  for (let ei = 0; ei < employees.length; ei++) {

    const name = employees[ei].name;



    let workDays = 0;

    for (let di = 0; di < dates.length; di++) {

      if (grid[ei][di] !== "休") workDays++;

    }



    if (!quotaRules.has(workDays)) {

      targetByName.set(name, null);

      warn.push([

        "-",

        `${name}：出勤日数${workDays}日の回数設定がないため、区分1/区分2の目標制御を行いません。`

      ]);

      continue;

    }



    targetByName.set(name, quotaRules.get(workDays));

  }



  return targetByName;

}



function categoryDeficit_(name, category, categoryTargetsByName, categoryCount) {

  if (!category) return 0;



  const targetMap = categoryTargetsByName.get(name);

  if (!targetMap) return 0;



  const target = targetMap.get(category) || 0;

  const current = categoryCount.get(name).get(category) || 0;



  return target - current;

}



function countPositiveCategoryDeficit_(names, category, categoryTargetsByName, categoryCount) {

  if (!category) return 0;



  let count = 0;

  for (const name of names) {

    if (categoryDeficit_(name, category, categoryTargetsByName, categoryCount) > 0) count++;

  }

  return count;

}



function categoryQuotaSortScore_(name, category, categoryTargetsByName, categoryCount) {

  if (!category) return 0;



  const deficit = categoryDeficit_(name, category, categoryTargetsByName, categoryCount);

  if (deficit > 0) return -deficit;



  return 1000 + Math.abs(deficit);

}



function eligibleCandidates_(

  job,

  availableEmployees,

  assignedToday,

  lastDone,

  lastConflictGroupDone,

  conflictGroupOfJob,

  dayIndex,

  intervalDays

) {

  const out = [];

  const g = conflictGroupOfJob.get(job);



  for (const name of availableEmployees) {

    if (assignedToday.has(name)) continue;



    const last = lastDone.get(name).get(job);

    if (typeof last === "number") {

      const diff = dayIndex - last;

      if (diff <= intervalDays) continue;

    }



    if (typeof g === "number") {

      const lastG = lastConflictGroupDone.get(name).get(g);

      if (typeof lastG === "number") {

        const diffG = dayIndex - lastG;

        if (diffG <= JOB_CONFLICT_MIN_GAP_DAYS) continue;

      }

    }



    out.push(name);

  }

  return out;

}



function enumerateDates_(start, end) {

  const s = new Date(start.getFullYear(), start.getMonth(), start.getDate());

  const e = new Date(end.getFullYear(), end.getMonth(), end.getDate());

  if (s.getTime() > e.getTime()) throw new Error("開始日が終了日より後です。");



  const dates = [];

  for (let d = new Date(s); d.getTime() <= e.getTime(); d.setDate(d.getDate() + 1)) {

    dates.push(new Date(d));

  }

  return dates;

}



function dateKey_(d) {

  return Utilities.formatDate(d, TZ, "yyyy-MM-dd");

}



function formatMD_(d) {

  return Utilities.formatDate(d, TZ, "M/d");

}



function parseMD_(text, yearRefDate) {

  const m = /^(\d{1,2})\s*\/\s*(\d{1,2})$/.exec(String(text || "").trim());

  if (!m) return null;

  const month = Number(m[1]);

  const day = Number(m[2]);

  if (!Number.isFinite(month) || !Number.isFinite(day)) return null;



  const y = yearRefDate.getFullYear();

  const dt = new Date(y, month - 1, day);

  if (dt.getFullYear() !== y || dt.getMonth() !== (month - 1) || dt.getDate() !== day) return null;

  return dt;

}



function normalizeToDate_(v, yearRefDate) {

  if (v instanceof Date && !isNaN(v.getTime())) {

    return new Date(v.getFullYear(), v.getMonth(), v.getDate());

  }

  const s = String(v || "").trim();

  if (!s) return null;



  const md = parseMD_(s, yearRefDate);

  if (md) return md;



  const dt = new Date(s);

  if (dt instanceof Date && !isNaN(dt.getTime())) {

    return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());

  }

  return null;

}



function createEmptyGrid_(rows, cols) {

  return Array.from({ length: rows }, () => Array(cols).fill(""));

}



function countBlankInGroup_(grid, di, idxArr) {

  let c = 0;

  for (const ei of idxArr) {

    if (grid[ei][di] === "") c++;

  }

  return c;

}



function countNonRestNonBlankInGroup_(grid, di, idxArr) {

  let c = 0;

  for (const ei of idxArr) {

    const v = grid[ei][di];

    if (v !== "" && v !== "休") c++;

  }

  return c;

}



function countUnexperienced_(names, job, jobCount) {

  let c = 0;

  for (const n of names) {

    if ((jobCount.get(n).get(job) || 0) === 0) c++;

  }

  return c;

}



function groupNeedPenalty_(name, employees, nameToIndex, mAssignedCount, mNeedAssign, aAssignedCount, aNeedAssign) {

  const ei = nameToIndex.get(name);

  const g = employees[ei].group;



  const mDef = mNeedAssign - mAssignedCount;

  const aDef = aNeedAssign - aAssignedCount;



  if (mDef > 0 && aDef <= 0) return (g === "午前") ? 0 : 1;

  if (aDef > 0 && mDef <= 0) return (g === "午後") ? 0 : 1;

  return 0;

}



function appendUnexperiencedWarnings_(warn, employees, jobs, grid, dates, jobCount, conflictGroupOfJob) {

  for (let ei = 0; ei < employees.length; ei++) {

    const name = employees[ei].name;

    const workDays = grid[ei].filter(v => v !== "休").length;



    const unexp = [];

    for (const job of jobs) {

      const c = (jobCount.get(name).get(job) || 0);

      if (c === 0) unexp.push(job);

    }

    if (unexp.length === 0) continue;



    if (workDays === 0) {

      warn.push(["-", `${name}：勤務可能日が0日のため未経験業務が残ります（未経験：${unexp.join(", ")}）`]);

      continue;

    }

    if (workDays < jobs.length) {

      warn.push(["-", `${name}：勤務可能日数${workDays} < 業務数${jobs.length} のため未経験業務が必ず残ります。`]);

    }



    const items = unexp.map(job => {

      const first = findFirstCandidateDayForJob_(ei, job, grid, dates, conflictGroupOfJob);

      return first ? `「${job}」(${first})` : `「${job}」(候補日なし)`;

    });



    const chunkSize = 5;

    for (let i = 0; i < items.length; i += chunkSize) {

      const part = items.slice(i, i + chunkSize).join(", ");

      warn.push(["-", `未経験業務（${name}）：${part}`]);

    }

  }

}



function findFirstCandidateDayForJob_(ei, targetJob, grid, dates, conflictGroupOfJob) {

  const gTarget = conflictGroupOfJob.get(targetJob);

  let lastSameGroupDay = null;



  for (let di = 0; di < dates.length; di++) {

    const cell = grid[ei][di];



    if (cell !== "休") {

      let ok = true;

      if (typeof gTarget === "number" && typeof lastSameGroupDay === "number") {

        if (di - lastSameGroupDay <= JOB_CONFLICT_MIN_GAP_DAYS) ok = false;

      }

      if (ok) return formatMD_(dates[di]);

    }



    if (cell && cell !== "休") {

      const gToday = conflictGroupOfJob.get(cell);

      if (typeof gTarget === "number" && gToday === gTarget) {

        lastSameGroupDay = di;

      }

    }

  }

  return null;

}



function appendCategoryQuotaWarnings_(warn, employees, categoryTargetsByName, categoryCount) {

  for (const e of employees) {

    const name = e.name;

    const targetMap = categoryTargetsByName.get(name);



    if (!targetMap) continue;



    for (const cat of COUNT_CATEGORY_NAMES) {

      const target = targetMap.get(cat) || 0;

      const actual = categoryCount.get(name).get(cat) || 0;



      if (actual !== target) {

        warn.push([

          "-",

          `${name}：${cat}の回数が目標と一致していません。目標=${target}回、実績=${actual}回、差=${actual - target}回`

        ]);

      }

    }

  }

}



function writeWarnings_(ss, warnRows) {

  const sh = ss.getSheetByName(SHEET_WARN) || ss.insertSheet(SHEET_WARN);

  sh.clear();

  sh.getRange(1, 1, 1, 2).setValues([["日付", "内容"]]);

  if (warnRows.length > 0) {

    sh.getRange(2, 1, warnRows.length, 2).setValues(warnRows);

    sh.autoResizeColumns(1, 2);

  }

}

/* =====================================================================
 * スマホ操作対応（既存機能・既存配置は変更しない追加部分）
 *
 * 設定シート D:E をスマホ操作欄として使用します。
 * 初回のみ Apps Script エディタから installMobileSupport() を実行してください。
 * ===================================================================== */

const MOBILE_CTRL_LABEL_COL = 4; // D列
const MOBILE_CTRL_VALUE_COL = 5; // E列

const MOBILE_ROW_TEMPLATE = 2;
const MOBILE_ROW_QUOTA = 3;
const MOBILE_ROW_FIX = 4;
const MOBILE_ROW_UNFIX = 5;
const MOBILE_ROW_GENERATE = 6;
const MOBILE_ROW_TARGET = 8;
const MOBILE_ROW_STATUS = 9;
const MOBILE_ROW_TIME = 10;

/**
 * 初回セットアップ。
 * 設定!D:E にスマホ操作欄を作成し、インストール型 onEdit トリガーを作成します。
 * この関数は最初に1回だけ Apps Script エディタから実行してください。
 */
function installMobileSupport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("スプレッドシートを取得できません。");

  setupMobileControlPanel_(ss);

  // 同じユーザーが作成した mobileOnEdit トリガーの重複を防止
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === "mobileOnEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger("mobileOnEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  const cfg = ss.getSheetByName(SHEET_CFG);
  cfg.getRange(MOBILE_ROW_STATUS, MOBILE_CTRL_VALUE_COL)
    .setValue("スマホ操作の初期設定が完了しました。");
  cfg.getRange(MOBILE_ROW_TIME, MOBILE_CTRL_VALUE_COL)
    .setValue(new Date())
    .setNumberFormat("yyyy/MM/dd HH:mm:ss");
}

/**
 * 設定シート D:E にスマホ操作欄を作成します。
 * A:C および既存の設定セルには触れません。
 */
function setupMobileControlPanel_(ss) {
  const cfg = ss.getSheetByName(SHEET_CFG);
  if (!cfg) throw new Error("設定シートが見つかりません。");

  cfg.getRange("D1:E10").clearContent().clearDataValidations();

  cfg.getRange("D1:E10").setValues([
    ["スマホ操作", "実行"],
    ["ひな形作成/更新", false],
    ["回数設定ひな形作成/更新", false],
    ["選択セルを固定", false],
    ["選択セルの固定を解除", false],
    ["生成", false],
    ["", ""],
    ["固定/解除の対象範囲", ""],
    ["実行結果", ""],
    ["実行日時", ""],
  ]);

  cfg.getRange("E2:E6").insertCheckboxes();
  cfg.getRange("E2:E6").setValue(false);

  cfg.getRange("D1:E1").setFontWeight("bold");
  cfg.getRange("D1:D10").setWrap(true);
  cfg.getRange("E8:E10").setWrap(true);
  cfg.setColumnWidth(MOBILE_CTRL_LABEL_COL, 210);
  cfg.setColumnWidth(MOBILE_CTRL_VALUE_COL, 180);
}

/**
 * iPhone / Android の Google スプレッドシートアプリから
 * 設定!E2:E6 のチェックボックスをONにしたときに実行されます。
 *
 * installMobileSupport() が作成する「インストール型編集トリガー」から呼ばれます。
 * この関数を手動実行しないでください。
 */
function mobileOnEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sh = range.getSheet();

  if (sh.getName() !== SHEET_CFG) return;
  if (range.getColumn() !== MOBILE_CTRL_VALUE_COL) return;

  const row = range.getRow();
  if (row < MOBILE_ROW_TEMPLATE || row > MOBILE_ROW_GENERATE) return;

  // チェックがONになったときだけ実行
  if (range.getValue() !== true) return;

  const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    range.setValue(false);
    setMobileStatus_(ss, "別の処理を実行中です。少し待ってからもう一度お試しください。", true);
    return;
  }

  try {
    setMobileStatus_(ss, "実行中...", false);

    let message = "";

    switch (row) {
      case MOBILE_ROW_TEMPLATE:
        mobilePrepareShiftTemplate_(ss);
        message = "ひな形作成/更新が完了しました。";
        break;

      case MOBILE_ROW_QUOTA:
        mobilePrepareQuotaTemplate_(ss);
        message = "回数設定ひな形作成/更新が完了しました。";
        break;

      case MOBILE_ROW_FIX: {
        const count = mobileMarkTargetCellsAsFixed_(ss);
        message = `${count}個のセルを固定しました。`;
        break;
      }

      case MOBILE_ROW_UNFIX: {
        const count = mobileUnmarkTargetCellsAsFixed_(ss);
        message = `${count}個のセルの固定を解除しました。`;
        break;
      }

      case MOBILE_ROW_GENERATE:
        generateShiftCore_(ss);
        message = "シフト生成が完了しました。";
        break;

      default:
        return;
    }

    setMobileStatus_(ss, message, false);

  } catch (err) {
    const message = err && err.message ? err.message : String(err);
    setMobileStatus_(ss, `エラー: ${message}`, true);

  } finally {
    // スクリプトによる変更では onEdit は再発火しないため、安全にOFFへ戻せます。
    range.setValue(false);

    const cfg = ss.getSheetByName(SHEET_CFG);
    if (cfg) {
      cfg.getRange(MOBILE_ROW_TIME, MOBILE_CTRL_VALUE_COL)
        .setValue(new Date())
        .setNumberFormat("yyyy/MM/dd HH:mm:ss");
    }

    lock.releaseLock();
  }
}

/**
 * シフト表で選択した範囲を設定!E8へ記録します。
 * モバイルアプリで選択変更イベントが取得できない場合は、
 * 設定!E8 に B2 や B2:D4 のように直接入力しても使用できます。
 */
function onSelectionChange(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_OUT) return;

  const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEET_CFG);
  if (!cfg) return;

  cfg.getRange(MOBILE_ROW_TARGET, MOBILE_CTRL_VALUE_COL)
    .setValue(e.range.getA1Notation());
}

/**
 * PC版「ひな形作成/更新」と同じ処理を、UIダイアログなしで実行します。
 * 元の prepareShiftTemplate() 自体は変更していません。
 */
function mobilePrepareShiftTemplate_(ss) {
  const employees = readEmployees_(ss);
  const cfg = readConfig_(ss);
  const dates = enumerateDates_(cfg.startDate, cfg.endDate);

  const fixedRead = readFixedEntriesFromOutput_(ss, employees, dates);
  const fixedByCell = new Map();

  for (const f of fixedRead.entries) {
    fixedByCell.set(cellKey_(f.ei, f.di), f.value);
  }

  const out = ss.getSheetByName(SHEET_OUT) || ss.insertSheet(SHEET_OUT);
  writeShiftOutput_(
    out,
    employees,
    dates,
    createEmptyGrid_(employees.length, dates.length),
    fixedByCell
  );
}

/**
 * PC版「回数設定ひな形作成/更新」と同じ処理を、
 * UIダイアログなしで実行します。
 * 元の prepareQuotaTemplate() 自体は変更していません。
 */
function mobilePrepareQuotaTemplate_(ss) {
  const sh = ss.getSheetByName(SHEET_QUOTA) || ss.insertSheet(SHEET_QUOTA);

  sh.getRange(1, 1, 1, 3).setValues([["出勤日数", "区分1", "区分2"]]);

  const maxRows = Math.max(1, sh.getMaxRows() - 1);
  const existingQuota = sh.getRange(2, 1, maxRows, 3).getValues()
    .some(row => row.some(v => String(v || "").trim() !== ""));

  if (!existingQuota) {
    sh.getRange(2, 1, 2, 3).setValues([
      [10, 5, 5],
      [12, 6, 6],
    ]);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 3);
}

/**
 * 設定!E8 に保存されたシフト表の範囲を取得します。
 * 「B2」「B2:D4」「B2:D4,F5」のような指定に対応します。
 */
function getMobileTargetRanges_(ss) {
  const cfg = ss.getSheetByName(SHEET_CFG);
  if (!cfg) throw new Error("設定シートが見つかりません。");

  const out = ss.getSheetByName(SHEET_OUT);
  if (!out) throw new Error("シフト表シートが見つかりません。");

  let target = String(
    cfg.getRange(MOBILE_ROW_TARGET, MOBILE_CTRL_VALUE_COL).getDisplayValue() || ""
  ).trim();

  if (!target) {
    throw new Error(
      "固定/解除の対象範囲がありません。シフト表で対象セルを選択するか、設定!E8 に B2 や B2:D4 のように入力してください。"
    );
  }

  // 「シフト表!B2:D4」「'シフト表'!B2:D4」にも対応
  target = target
    .replace(/^'シフト表'!/i, "")
    .replace(/^シフト表!/i, "");

  const a1List = target
    .split(/[,\n、]+/)
    .map(s => s.trim())
    .filter(Boolean);

  if (a1List.length === 0) {
    throw new Error("固定/解除の対象範囲が不正です。");
  }

  const ranges = [];
  for (const a1 of a1List) {
    try {
      ranges.push(out.getRange(a1));
    } catch (err) {
      throw new Error(`対象範囲「${a1}」を読み取れません。B2 や B2:D4 の形式で指定してください。`);
    }
  }

  return ranges;
}

/**
 * スマホ用：対象範囲を固定します。
 * 固定判定・色・ノートは既存仕様（FIXED_NOTE / FIXED_BG）と同一です。
 */
function mobileMarkTargetCellsAsFixed_(ss) {
  const ranges = getMobileTargetRanges_(ss);
  let count = 0;

  for (const range of ranges) {
    const values = range.getValues();
    const notes = range.getNotes();
    const backgrounds = range.getBackgrounds();

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[0].length; c++) {
        const row = range.getRow() + r;
        const col = range.getColumn() + c;

        if (row < 2 || col < 2) continue;

        const value = normalizeCellValue_(values[r][c]);
        if (!value) continue;

        notes[r][c] = FIXED_NOTE;
        backgrounds[r][c] = FIXED_BG;
        count++;
      }
    }

    range.setNotes(notes);
    range.setBackgrounds(backgrounds);
  }

  return count;
}

/**
 * スマホ用：対象範囲の固定を解除します。
 * 解除判定・背景色は既存仕様と同一です。
 */
function mobileUnmarkTargetCellsAsFixed_(ss) {
  const ranges = getMobileTargetRanges_(ss);
  let count = 0;

  for (const range of ranges) {
    const notes = range.getNotes();
    const backgrounds = range.getBackgrounds();

    for (let r = 0; r < notes.length; r++) {
      for (let c = 0; c < notes[0].length; c++) {
        const row = range.getRow() + r;
        const col = range.getColumn() + c;

        if (row < 2 || col < 2) continue;

        if (isFixedMark_(notes[r][c], backgrounds[r][c])) {
          notes[r][c] = "";
          backgrounds[r][c] = AUTO_BG;
          count++;
        }
      }
    }

    range.setNotes(notes);
    range.setBackgrounds(backgrounds);
  }

  return count;
}

/**
 * 設定!E9 に実行結果を表示します。
 */
function setMobileStatus_(ss, message, isError) {
  const cfg = ss.getSheetByName(SHEET_CFG);
  if (!cfg) return;

  const cell = cfg.getRange(MOBILE_ROW_STATUS, MOBILE_CTRL_VALUE_COL);
  cell.setValue(message);

  // 色は付けず、エラー時のみ太字にして区別
  cell.setFontWeight(isError ? "bold" : "normal");
}
