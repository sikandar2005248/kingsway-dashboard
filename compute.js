/* Works out every dashboard figure from the workbook's sheets.
   Same rules as the Kingsway Finance workbook:
   - balances = opening (Settings) + income - expenses marked Paid + transfers in - transfers out
   - undated BSN archive receipts are placed at their semester's estimated start. */
(function () {
  'use strict';

  const ACCOUNTS = ['Cash', 'Bank', 'Easypaisa'];

  const num = (v) => {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    const n = Number(String(v).replace(/,/g, '').trim());
    return isNaN(n) ? 0 : n;
  };
  const text = (v) => (v === null || v === undefined ? '' : String(v).trim());
  const pad = (n) => String(n).padStart(2, '0');
  const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

  function iso(y, m, d) {
    const t = new Date(Date.UTC(y, m - 1, d));
    if (isNaN(t) || t.getUTCMonth() !== m - 1) return '';
    return t.getUTCFullYear() + '-' + pad(t.getUTCMonth() + 1) + '-' + pad(t.getUTCDate());
  }

  function toDate(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number') {
      if (v < 1 || v > 2958465) return '';
      const t = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
      return t.getUTCFullYear() + '-' + pad(t.getUTCMonth() + 1) + '-' + pad(t.getUTCDate());
    }
    const s = String(v).trim();
    let m;
    if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return iso(+m[1], +m[2], +m[3]);
    if ((m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/))) return iso(+m[3], +m[2], +m[1]);
    if ((m = s.match(/^(\d{1,2})[\s\-]([A-Za-z]{3})[A-Za-z]*[\s\-,]+(\d{4})$/)) && MONTHS[m[2].toLowerCase()]) return iso(+m[3], MONTHS[m[2].toLowerCase()], +m[1]);
    return '';
  }

  // Sheet rows as objects. Header row = first of the top 10 rows containing `mustHave`.
  function rowsOf(book, sheet, mustHave) {
    const rows = book[sheet];
    if (!rows) return [];
    let h = -1;
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      if ((rows[i] || []).some((c) => text(c) === mustHave)) { h = i; break; }
    }
    if (h < 0) return [];
    const header = (rows[h] || []).map(text);
    const out = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      if (!r.some((c) => c !== null && c !== undefined && text(c) !== '')) continue;
      const o = {};
      header.forEach((k, j) => { if (k) o[k] = r[j] === undefined ? null : r[j]; });
      out.push(o);
    }
    return out;
  }

  function addMonths(y, m, n) { const k = (m - 1) + n; return [y + Math.floor(k / 12), ((k % 12) + 12) % 12 + 1]; }
  function weekOf(d) {
    const t = new Date(d + 'T00:00:00Z');
    t.setUTCDate(t.getUTCDate() - ((t.getUTCDay() + 6) % 7));
    return t.toISOString().slice(0, 10);
  }
  const sum = (arr, f) => arr.reduce((a, r) => a + f(r), 0);

  function compute(book, semesterOverrides) {
    semesterOverrides = semesterOverrides || {};
    const raw = {
      bsn: rowsOf(book, 'BSN Fee Collection', 'Amount Collected (PKR)'),
      archive: rowsOf(book, 'BSN Archive Fee Collection', 'Amount Collected (PKR)'),
      catb: rowsOf(book, 'Category B Fees', 'Fee Collected (PKR)'),
      fines: rowsOf(book, 'Fine Income', 'Fine Amount (PKR)'),
      expenses: rowsOf(book, 'Expense Tracker', 'Amount (PKR)'),
      transfers: rowsOf(book, 'Account Transfers', 'Amount (PKR)'),
      settings: rowsOf(book, 'Settings', 'Financial Medium'),
      challans: rowsOf(book, 'Fee Challans', 'Challan No.'),
      students: rowsOf(book, 'Students', 'Registration No.'),
    };
    if (!raw.bsn.length && !raw.expenses.length && !raw.settings.length) {
      throw new Error("This doesn't look like the Kingsway Finance workbook (no BSN Fee Collection, Expense Tracker or Settings sheet found).");
    }

    // ---- income
    const income = [];
    raw.bsn.forEach((r) => income.push({
      date: toDate(r['Collection Date']), source: 'BSN ' + (text(r['Fee Type']) || 'Fee'),
      receipt: text(r['Receipt No.']), reg: text(r['Student ID / Reg. No.']), student: text(r['Student Name']),
      batch: text(r['Batch']), semester: text(r['Semester']), account: text(r['Payment Medium']),
      discount: num(r['Discount (PKR)']), amount: num(r['Amount Collected (PKR)']), remarks: text(r['Description / Remarks']),
    }));
    raw.catb.forEach((r) => income.push({
      date: toDate(r['Collection Date']), source: 'Category B ' + (text(r['Fee Type']) || 'Fee'),
      receipt: text(r['Receipt No.']), reg: text(r['Student ID / Reg. No.']), student: text(r['Student Name']),
      batch: text(r['Batch']), semester: '', account: text(r['Payment Medium']),
      discount: 0, amount: num(r['Fee Collected (PKR)']), remarks: text(r['Description / Remarks']),
    }));
    raw.fines.forEach((r) => income.push({
      date: toDate(r['Collection Date']), source: ('Fine ' + text(r['Fine Type'])).trim(),
      receipt: text(r['Receipt No.']), reg: text(r['Student ID / Reg. No.']), student: text(r['Student Name']),
      batch: text(r['Batch']), semester: '', account: text(r['Payment Medium']),
      discount: 0, amount: num(r['Fine Amount (PKR)']), remarks: text(r['Description / Remarks']),
    }));
    const dated = income.filter((r) => r.date);
    const undated = income.filter((r) => !r.date);

    // ---- expenses, transfers, openings
    const expenses = raw.expenses.map((r) => ({
      date: toDate(r['Expense Date']), category: text(r['Category']), description: text(r['Description']),
      vendor: text(r['Paid To / Vendor']), amount: num(r['Amount (PKR)']), account: text(r['Payment Medium']),
      status: text(r['Payment Status']),
    }));
    const paid = expenses.filter((e) => e.status.toLowerCase() === 'paid');
    const transfers = raw.transfers.map((r) => ({
      date: toDate(r['Transfer Date']), from: text(r['From Medium']), to: text(r['To Medium']),
      amount: num(r['Amount (PKR)']), remarks: text(r['Description / Remarks']),
    }));
    const openings = {};
    ACCOUNTS.forEach((a) => { openings[a] = 0; });
    raw.settings.forEach((r) => { const m = text(r['Financial Medium']); if (m) openings[m] = num(r['Opening Balance (PKR)']); });
    const accounts = Object.keys(openings);

    // ---- ledger
    const events = [];
    income.forEach((r) => events.push({ date: r.date, account: r.account, ord: 1, type: 'Income',
      detail: r.source + (r.student ? ' · ' + r.student : ''), money_in: r.amount, money_out: 0 }));
    transfers.forEach((t) => {
      events.push({ date: t.date, account: t.from, ord: 2, type: 'Transfer out', detail: 'To ' + t.to + (t.remarks ? ' · ' + t.remarks : ''), money_in: 0, money_out: t.amount });
      events.push({ date: t.date, account: t.to, ord: 2, type: 'Transfer in', detail: 'From ' + t.from + (t.remarks ? ' · ' + t.remarks : ''), money_in: t.amount, money_out: 0 });
    });
    paid.forEach((e) => events.push({ date: e.date, account: e.account, ord: 3, type: 'Expense',
      detail: [e.category, e.vendor, e.description].filter(Boolean).join(' · '), money_in: 0, money_out: e.amount }));
    events.forEach((e) => { if (e.account && !(e.account in openings)) { openings[e.account] = 0; accounts.push(e.account); } });

    const ledger = [];
    accounts.forEach((a) => {
      let bal = openings[a];
      ledger.push({ id: 'open-' + a, date: '', account: a, type: 'Opening balance', detail: 'From the Settings sheet', money_in: 0, money_out: 0, balance: bal });
      events.filter((e) => e.account === a)
        .sort((x, y) => (x.date < y.date ? -1 : x.date > y.date ? 1 : x.ord - y.ord))
        .forEach((e, i) => { bal += e.money_in - e.money_out; ledger.push({ id: a + '-' + i, date: e.date, account: a, type: e.type, detail: e.detail, money_in: e.money_in, money_out: e.money_out, balance: bal }); });
    });

    // ---- balances
    const balances = accounts.map((a) => {
      const ev = events.filter((e) => e.account === a);
      const inc = sum(ev.filter((e) => e.type === 'Income'), (e) => e.money_in);
      const exp = sum(ev.filter((e) => e.type === 'Expense'), (e) => e.money_out);
      const tin = sum(ev.filter((e) => e.type === 'Transfer in'), (e) => e.money_in);
      const tout = sum(ev.filter((e) => e.type === 'Transfer out'), (e) => e.money_out);
      return { account: a, opening: openings[a], income: inc, expenses: exp, transfers_in: tin, transfers_out: tout, closing: openings[a] + inc - exp + tin - tout };
    });
    const total = { account: 'Total' };
    ['opening', 'income', 'expenses', 'transfers_in', 'transfers_out', 'closing'].forEach((k) => { total[k] = sum(balances, (b) => b[k]); });

    // ---- opening / closing by day
    const dates = Array.from(new Set(events.map((e) => e.date).filter(Boolean))).sort();
    const daily_balances = [];
    accounts.forEach((a) => {
      let bal = openings[a];
      dates.forEach((d) => {
        const day = events.filter((e) => e.account === a && e.date === d);
        const mi = sum(day, (e) => e.money_in), mo = sum(day, (e) => e.money_out);
        daily_balances.push({ date: d, account: a, opening: bal, money_in: mi, money_out: mo, closing: bal + mi - mo });
        bal += mi - mo;
      });
    });

    // ---- semester starts
    const earliest = {};
    raw.students.forEach((s) => {
      const b = text(s['Batch']), d = toDate(s['Admission Date']);
      if (b && d && (!earliest[b] || d < earliest[b])) earliest[b] = d;
    });
    const starts = {};
    const semester_starts = [];
    const batchKeys = Array.from(new Set(Object.keys(earliest).concat(Object.keys(semesterOverrides))))
      .sort((x, y) => (x.length - y.length) || (x < y ? -1 : 1));
    batchKeys.forEach((b) => {
      let y, m, basis;
      if (semesterOverrides[b]) { [y, m] = String(semesterOverrides[b]).split('-').map(Number); basis = 'Set in config.js'; }
      else if (earliest[b]) { y = +earliest[b].slice(0, 4); m = +earliest[b].slice(5, 7); basis = 'Earliest admission ' + earliest[b]; }
      else return;
      starts[b] = [y, m];
      semester_starts.push({ batch: b, start_month: y + '-' + pad(m), basis });
    });

    const archive = raw.archive.map((r) => ({ batch: text(r['Batch']), semester: text(r['Semester']), date: toDate(r['Collection Date']), amount: num(r['Amount Collected (PKR)']) }));
    undated.forEach((r) => archive.push({ batch: r.batch, semester: r.semester || '1', date: '', amount: r.amount }));

    // ---- income by period
    const buckets = { daily: {}, weekly: {}, monthly: {}, yearly: {} };
    const add = (g, p, i, v) => { const b = buckets[g][p] || (buckets[g][p] = [0, 0]); b[i] += v; };
    const addDated = (d, v) => { add('daily', d, 0, v); add('weekly', weekOf(d), 0, v); add('monthly', d.slice(0, 7), 0, v); add('yearly', d.slice(0, 4), 0, v); };
    dated.forEach((r) => addDated(r.date, r.amount));
    let unplaced = 0;
    archive.forEach((r) => {
      if (r.date) { addDated(r.date, r.amount); return; }
      const st = starts[r.batch];
      if (!st) { unplaced += r.amount; return; }
      const sem = Math.max(1, parseInt(parseFloat(r.semester || '1'), 10) || 1);
      const [y, m] = addMonths(st[0], st[1], 6 * (sem - 1));
      add('monthly', y + '-' + pad(m), 1, r.amount);
      add('yearly', String(y), 1, r.amount);
    });
    const mk = Object.keys(buckets.monthly).sort();
    if (mk.length) {
      let [y, m] = mk[0].split('-').map(Number);
      const [ey, em] = mk[mk.length - 1].split('-').map(Number);
      while (y < ey || (y === ey && m <= em)) { const k = y + '-' + pad(m); if (!buckets.monthly[k]) buckets.monthly[k] = [0, 0]; [y, m] = addMonths(y, m, 1); }
    }
    const periods = {};
    Object.keys(buckets).forEach((g) => {
      periods[g] = Object.keys(buckets[g]).sort().map((p) => ({ period: p, recorded: buckets[g][p][0], estimated: buckets[g][p][1], total: buckets[g][p][0] + buckets[g][p][1] }));
    });
    const recorded_total = sum(dated, (r) => r.amount) + sum(archive.filter((r) => r.date), (r) => r.amount);
    const estimated_total = sum(archive.filter((r) => !r.date), (r) => r.amount) - unplaced;

    // ---- breakdowns
    const group = (rows, key) => {
      const m = {};
      rows.forEach((r) => { const k = r[key] || '(blank)'; m[k] = (m[k] || 0) + r.amount; });
      return Object.keys(m).map((k) => ({ label: k, amount: m[k] })).sort((a, b) => b.amount - a.amount);
    };
    const dayKeys = Array.from(new Set(dated.map((r) => r.date).concat(paid.map((e) => e.date).filter(Boolean)))).sort();
    const daily = dayKeys.map((d) => ({ date: d, income: sum(dated.filter((r) => r.date === d), (r) => r.amount), expenses: sum(paid.filter((e) => e.date === d), (e) => e.amount) }));

    const ab = {};
    raw.archive.forEach((r) => {
      const k = text(r['Batch']) + '|' + text(r['Semester']);
      const v = ab[k] || (ab[k] = { b: text(r['Batch']), s: text(r['Semester']), amount: 0, receipts: 0 });
      v.amount += num(r['Amount Collected (PKR)']); v.receipts += 1;
    });
    const ik = (s) => { const n = parseInt(s, 10); return isNaN(n) ? 999 : n; };
    const archive_by_batch = Object.values(ab).sort((x, y) => ik(x.b) - ik(y.b) || ik(x.s) - ik(y.s))
      .map((v) => ({ label: 'Batch ' + v.b + ' · Sem ' + v.s, amount: v.amount, receipts: v.receipts }));

    const unverified = raw.challans.filter((r) => text(r['Status']) === 'Paid-Unverified').map((r) => ({
      challan: text(r['Challan No.']), student: text(r['Student Name']), reg: text(r['Registration No.']),
      ref: text(r['Bank / Transaction Ref']), amount: num(r['Total (PKR)']),
    }));

    const dd = dated.map((r) => r.date).sort();
    const incTotal = sum(dated, (r) => r.amount), expTotal = sum(paid, (e) => e.amount);
    return {
      period: { first: dd[0] || '', last: dd[dd.length - 1] || '' },
      kpis: {
        income: incTotal, expenses: expTotal, net: incTotal - expTotal, closing: total.closing,
        discounts: sum(dated, (r) => r.discount), unverified: sum(unverified, (u) => u.amount), unverified_count: unverified.length,
      },
      balances, balances_total: total, daily_balances, ledger, periods,
      all_time: { recorded: recorded_total, estimated: estimated_total, total: recorded_total + estimated_total, unplaced },
      semester_starts, daily,
      by_source: group(dated, 'source'),
      by_category: group(paid.map((e) => ({ category: e.category, amount: e.amount })), 'category'),
      expenses: expenses.slice().sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0)),
      income: dated.slice().sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0)),
      unverified, archive_by_batch,
      archive_count: raw.archive.length,
      archive_total: sum(raw.archive, (r) => num(r['Amount Collected (PKR)'])),
    };
  }

  window.kwCompute = compute;
})();
