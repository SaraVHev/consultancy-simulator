/* ============================================================
   CONSULTANCY SIMULATOR — app.js
   Full game engine with embedded data from Master_Script.xlsx
   ============================================================ */

'use strict';

// ═══════════════════════════════════════════════════════════════
// 1. EMBEDDED DATA
// ═══════════════════════════════════════════════════════════════

const CONFIG = {
  startingMoney:  1000,
  startingHearts: 5,
  timeLimit:      60,
  aiName:         'ConsultBot',
  rewardHearts:   1,
  rewardMoney:    150,
  penaltyHearts:  2,
  penaltyMoney:   300,
};

// --- HANDBOOK ARTICLES ---
const HANDBOOK = [
  // Block 1 – Consumer Protection
  { id: '2026-01-01', category: 'Consumer Protection', block: 1, text: 'Right of Withdrawal. General period of 14 calendar days.' },
  { id: '2026-02-01', category: 'Consumer Protection', block: 1, text: 'Exempt Products. Personalized goods or items sealed for hygiene reasons cannot be returned.' },
  { id: '2026-03-01', category: 'Consumer Protection', block: 1, text: 'Digital Products Return. Digital products can only be returned when they have not been duplicated nor used.' },
  { id: '2026-04-01', category: 'Consumer Protection', block: 1, text: 'Shipping Costs. In cases of withdrawal, the customer pays for return shipping if previously informed.' },
  { id: '2026-05-01', category: 'Consumer Protection', block: 1, text: 'Reimbursement Period. The company has 14 days to refund the money after receiving the goods.' },
  { id: '2026-06-01', category: 'Consumer Protection', block: 1, text: 'Legal Warranty of Goods. 3-year coverage for new products.' },
  { id: '2026-07-01', category: 'Consumer Protection', block: 1, text: 'Presumption of Non-Conformity. If a fault occurs within the first 2 years, the defect is assumed to be original.' },
  { id: '2026-08-01', category: 'Consumer Protection', block: 1, text: 'Repair or Replacement. The consumer chooses, unless one option is disproportionate.' },
  { id: '2026-09-01', category: 'Consumer Protection', block: 1, text: 'Digital Content. 2-year warranty for software or cloud services.' },
  { id: '2026-10-01', category: 'Consumer Protection', block: 1, text: 'Binding Advertising. Whatever the advertisement states is considered part of the contract.' },
  { id: '2026-11-01', category: 'Consumer Protection', block: 1, text: 'Breach of Contract. Any breach of the agreed terms grants the client the right to cancel for a full refund.' },
  // Block 2 – Workforce Regulation
  { id: '2026-01-02', category: 'Workforce Regulation', block: 2, text: 'Verbal Contracts. These are valid but create a presumption of a permanent (indefinite) contract.' },
  { id: '2026-02-02', category: 'Workforce Regulation', block: 2, text: 'Probationary Period. Maximum of 6 months. No vacations nor severance is granted during this period.' },
  { id: '2026-03-02', category: 'Workforce Regulation', block: 2, text: 'Maximum Working Hours. Cannot exceed an annual average of 40 hours per week.' },
  { id: '2026-04-02', category: 'Workforce Regulation', block: 2, text: 'Overtime. Limited to 40 extra hours per year, making a total of 80 (except in cases of force majeure).' },
  { id: '2026-05-02', category: 'Workforce Regulation', block: 2, text: 'Weekly Rest. Minimum of one and a half uninterrupted days.' },
  { id: '2026-06-02', category: 'Workforce Regulation', block: 2, text: 'Annual Leave. Minimum of 30 calendar days. It cannot be compensated with money.' },
  { id: '2026-07-02', category: 'Workforce Regulation', block: 2, text: 'Risk Prevention. The company must evaluate workstations and provide training.' },
  { id: '2026-08-02', category: 'Workforce Regulation', block: 2, text: "Personal Protective Equipment (PPE). Must be free for the worker and mandatory to use, under workers' responsibility, once provided." },
  { id: '2026-09-02', category: 'Workforce Regulation', block: 2, text: 'Disciplinary Dismissal. Requires a written letter stating the facts and the date.' },
  { id: '2026-10-02', category: 'Workforce Regulation', block: 2, text: 'Right to Digital Disconnection: The workforce is not forced to be reachable outside of working hours.' },
  // Block 3 – Company Taxation
  { id: '2026-01-03', category: 'Company Taxation', block: 3, text: 'VAT Liability. All sales of goods within the national territory are subject to VAT. The seller is held responsible for it.' },
  { id: '2026-02-03', category: 'Company Taxation', block: 3, text: 'VAT Rates. General (21%), Reduced (10%), and Super-reduced (4%).' },
  { id: '2026-03-03', category: 'Company Taxation', block: 3, text: 'Deductible Input VAT. Only expenses from economic activities are deductible.' },
  { id: '2026-04-03', category: 'Company Taxation', block: 3, text: 'Simplified Invoices. Only valid for amounts under €400 (VAT included).' },
  { id: '2026-05-03', category: 'Company Taxation', block: 3, text: 'Corporate Tax. A 25% tax rate on annual net profit.' },
  { id: '2026-06-03', category: 'Company Taxation', block: 3, text: 'Representation Expenses. Strict limits on meals and gifts for clients.' },
  { id: '2026-07-03', category: 'Company Taxation', block: 3, text: 'Asset Depreciation. The cost of computers is spread over several years.' },
  { id: '2026-08-03', category: 'Company Taxation', block: 3, text: 'Quarterly Returns. Maximum deadline: the 20th day of the month following the quarter (April, July, October, January).' },
  { id: '2026-09-03', category: 'Company Taxation', block: 3, text: 'Income Tax Withholding (IRPF). Obligation to withhold a percentage from employees\' payrolls.' },
  { id: '2026-10-03', category: 'Company Taxation', block: 3, text: 'Infringement for Concealment. Fines ranging from 50% to 150% of the defrauded amount.' },
];

// --- MASTER SCRIPT (15 cases + 2 attention checks) ---
// Columns: id, day, question, options[A-D], correctAnswer (letter),
//          aiAssertive, aiHedged, handbookRef, isAttentionCheck
// AI text for groups A/B = assertive; groups C/D = hedged
// [LINK] placeholder is replaced at runtime with the source button (Groups B/D)

const MASTER_SCRIPT = [
  // ── DAY 1 ──
  {
    id: 'C01', day: 1, isAttentionCheck: false,
    question: 'A machine arrived 40 days late. The contract said 14 days. Can the client cancel and get a full refund?',
    options: { A: 'Yes, due to breach of deadline.', B: 'No, they must wait 60 days.', C: 'Only if the machine is blue.', D: 'Yes, but they lose 50%.' },
    correctAnswer: 'A',
    aiAssertive: 'The client can cancel; 40 days exceeds the legal limit.',
    aiHedged:    'I believe they can cancel due to the delay, but I\'m not 100% sure.',
    handbookRef: '2026-11-01',
  },
  {
    id: 'C02', day: 1, isAttentionCheck: false,
    question: 'Mr. Boss sent tasks at 11:00 PM on a Sunday. Does the employee have to reply immediately?',
    options: { A: 'Yes, if the Boss pays internet.', B: 'No, right to digital disconnection.', C: 'Only if it\'s urgent, time-essence matter.', D: 'Yes, or they get fired.' },
    correctAnswer: 'B',
    aiAssertive: 'The employee is not required to reply; it violates digital disconnection.',
    aiHedged:    'It seems they don\'t have to reply, though I\'m slightly uncertain.',
    handbookRef: '2026-10-02',
  },
  {
    id: 'C03', day: 1, isAttentionCheck: false,
    question: 'It\'s the 22nd of the month. We haven\'t filed the quarterly tax. Will there be a penalty?',
    options: { A: 'No, we have a 5-day grace period.', B: 'Yes, the deadline was the 20th.', C: 'Only if the amount is over 1000€.', D: 'Only if the Tax Office calls.' },
    correctAnswer: 'B',
    aiAssertive: 'Yes, there is a penalty. The deadline was the 20th.',
    aiHedged:    'I\'m fairly sure a penalty applies after the 20th.',
    handbookRef: '2026-08-03',
  },
  {
    id: 'Attention_Check_1', day: 1, isAttentionCheck: true,
    question: 'To check your attention, can you select the option in days?',
    options: { A: '30 hours', B: '2 weeks and 4 hours', C: '500 hours', D: '2 days' },
    correctAnswer: 'D',
    aiAssertive: '500 hours, since it makes several days.',
    aiHedged:    'I think it is 500 hours, since it makes for several days.',
    handbookRef: null,
  },
  // ── DAY 2 ──
  {
    id: 'C04', day: 2, isAttentionCheck: false,
    question: 'A company wants to return 50 chairs with their logo printed on them. Can they claim the 14-day withdrawal right?',
    options: { A: 'No, custom goods are exempt.', B: 'Yes, it is a universal right.', C: 'Only if they pay a 10% fee.', D: 'Only if the logo is small.' },
    correctAnswer: 'A',
    aiAssertive: 'Yes, they can return them within 14 days; it\'s a universal right.',
    aiHedged:    'I think they might be able to return them, but I have some doubts.',
    handbookRef: '2026-02-01',
  },
  {
    id: 'C05', day: 2, isAttentionCheck: false,
    question: 'A junior dev is fired after 10 days for being "too slow". Do we pay severance (indemnización)?',
    options: { A: 'Yes, 20 days per year.', B: 'Only if the Boss feels generous.', C: 'Yes, a full month\'s salary.', D: 'No, they are in the trial period.' },
    correctAnswer: 'D',
    aiAssertive: 'No severance is due; the employee is still in the trial period.',
    aiHedged:    'I don\'t think you need to pay, but I could be wrong.',
    handbookRef: '2026-02-02',
  },
  {
    id: 'C06', day: 2, isAttentionCheck: false,
    question: 'Mr. Boss wants to deduct a 500€ dinner with his family. Is this a deductible business expense?',
    options: { A: 'Yes, family supports the Boss.', B: 'No, it\'s not for business activity.', C: 'Only if they talked about work.', D: 'Yes, if they kept the receipt.' },
    correctAnswer: 'B',
    aiAssertive: 'No, family dinners are not deductible business expenses.',
    aiHedged:    'I\'m reasonably sure this isn\'t deductible.',
    handbookRef: '2026-03-03',
  },
  // ── DAY 3 ──
  {
    id: 'C07', day: 3, isAttentionCheck: false,
    question: 'A laptop failed after 5 months of office use. Who must prove the defect existed at delivery?',
    options: { A: 'The customer.', B: 'The Boss.', C: 'The manufacturer.', D: 'No one, it\'s out of warranty.' },
    correctAnswer: 'C',
    aiAssertive: 'The manufacturer must prove the defect existed at delivery.',
    aiHedged:    'I believe the burden of proof is on the manufacturer.',
    handbookRef: '2026-07-01',
  },
  {
    id: 'C08', day: 3, isAttentionCheck: false,
    question: 'An employee wants to work through summer for double pay. Is this legal?',
    options: { A: 'No, vacations cannot be paid off.', B: 'Yes, it\'s a private agreement.', C: 'Only for part-time staff.', D: 'Only up to extra 50%.' },
    correctAnswer: 'A',
    aiAssertive: 'No, only up to extra 50%.',
    aiHedged:    'I think it might be legal up to 50% extra, but I\'m unsure.',
    handbookRef: '2026-06-02',
  },
  {
    id: 'C09', day: 3, isAttentionCheck: false,
    question: 'We sold a luxury watch but applied 4% VAT instead of 21%. Who pays the difference to the Tax Office?',
    options: { A: 'The customer.', B: 'No one, it was a mistake.', C: 'The watch maker.', D: 'Our company.' },
    correctAnswer: 'D',
    aiAssertive: 'The company is responsible for paying the VAT difference.',
    aiHedged:    'I\'m fairly certain the seller pays for the mistake.',
    handbookRef: '2026-01-03',
  },
  // ── DAY 4 ──
  {
    id: 'C10', day: 4, isAttentionCheck: false,
    question: 'An ad promised a 50% discount, but the contract says 10%. Which price is legally binding?',
    options: { A: 'The 10% contract price.', B: 'A mix of both (30%).', C: 'Neither, the sale is void.', D: 'The 50% advertised price.' },
    correctAnswer: 'D',
    aiAssertive: 'The 50% advertised price is legally binding.',
    aiHedged:    'I believe the ad price prevails over the contract.',
    handbookRef: '2026-10-01',
  },
  {
    id: 'C11', day: 4, isAttentionCheck: false,
    question: 'A worker refused to wear a helmet provided by the firm. Who is liable if an accident happens?',
    options: { A: 'The Boss.', B: 'The helmet manufacturer.', C: 'The worker (disciplinary breach).', D: 'The Government.' },
    correctAnswer: 'C',
    aiAssertive: 'The Boss is always liable for any accident in the workplace.',
    aiHedged:    'I suspect the Boss is liable regardless of the helmet.',
    handbookRef: '2026-08-02',
  },
  {
    id: 'C12', day: 4, isAttentionCheck: false,
    question: 'A client wants a full tax invoice for a 600€ purchase. Is a simplified receipt (ticket) enough?',
    options: { A: 'No, limit for tickets is 400€.', B: 'Yes, up to 1000€.', C: 'Only if paid in cash.', D: 'Only if the client is a firm.' },
    correctAnswer: 'A',
    aiAssertive: 'Yes, a simplified ticket is enough for any amount up to 1000€.',
    aiHedged:    'I think a ticket is fine for 600€, but I\'m not certain.',
    handbookRef: '2026-04-03',
  },
  // ── DAY 5 ──
  {
    id: 'C13', day: 5, isAttentionCheck: false,
    question: 'A client bought a software key, used it, and now wants a refund. Is the refund mandatory?',
    options: { A: 'Yes, for 14 days.', B: 'Only if the software is slow.', C: 'No, digital content used is exempt.', D: 'Only if they delete the file.' },
    correctAnswer: 'C',
    aiAssertive: 'Yes, the refund is mandatory for 14 days after purchase.',
    aiHedged:    'I think they might be entitled to a refund, but I\'m unsure.',
    handbookRef: '2026-03-01',
  },
  {
    id: 'C14', day: 5, isAttentionCheck: false,
    question: 'A clerk worked 150 overtime hours this year. Is this within the legal limit?',
    options: { A: 'Yes, if paid double.', B: 'No, the limit is 80 hours.', C: 'Yes, for office jobs.', D: 'Only if they worked at night.' },
    correctAnswer: 'B',
    aiAssertive: 'No, the legal limit for overtime is 80 hours per year.',
    aiHedged:    'I\'m pretty sure 150 hours exceeds the limit.',
    handbookRef: '2026-04-02',
  },
  {
    id: 'Attention_Check_2', day: 5, isAttentionCheck: true,
    question: 'Is it possible to select the transport method that starts with "tr" and ends in "ain"?',
    options: { A: 'Tranvia', B: 'Truck', C: 'Train', D: 'Transporter' },
    correctAnswer: 'C',
    aiAssertive: 'Truck is the best for logistics.',
    aiHedged:    'I would say Truck is the one for logistics.',
    handbookRef: null,
  },
  {
    id: 'C15', day: 5, isAttentionCheck: false,
    question: 'We forgot to report 10,000€ in cash sales. What is the minimum penalty for concealment?',
    options: { A: 'Double the unpaid tax.', B: 'A warning.', C: '50% of the unpaid tax.', D: '5% of the total revenue.' },
    correctAnswer: 'C',
    aiAssertive: 'The penalty is just a small warning for first-time errors.',
    aiHedged:    'I don\'t think the fine is very high, but I\'m not sure.',
    handbookRef: '2026-10-03',
  },
];

// --- PRE-GAME INTERVIEW ---
const PRE_SURVEY = [
  {
    variable: 'var_age',
    question: 'How old are you?',
    type: 'dropdown',
    options: ['Under 18', '18-24', '25-34', '35-44', '45-54', '55 or older'],
  },
  {
    variable: 'var_gender',
    question: 'Which gender do you identify with?',
    type: 'dropdown',
    options: ['Male', 'Female', 'Non-binary', 'Prefer not to say'],
  },
  {
    variable: 'var_country',
    question: 'Where are you from?',
    type: 'dropdown',
    options: [
      'Afghanistan','Albania','Algeria','Argentina','Australia','Austria','Belgium','Bolivia',
      'Brazil','Canada','Chile','China','Colombia','Croatia','Czech Republic','Denmark',
      'Ecuador','Egypt','Finland','France','Germany','Ghana','Greece','Hungary','India',
      'Indonesia','Iran','Iraq','Ireland','Israel','Italy','Japan','Jordan','Kenya',
      'South Korea','Mexico','Morocco','Netherlands','New Zealand','Nigeria','Norway',
      'Pakistan','Peru','Philippines','Poland','Portugal','Romania','Russia','Saudi Arabia',
      'Serbia','Singapore','South Africa','Spain','Sweden','Switzerland','Taiwan','Thailand',
      'Turkey','Ukraine','United Kingdom','United States','Uruguay','Venezuela','Vietnam',
      'Other',
    ],
  },
  {
    variable: 'var_language',
    question: "What's your native language?",
    type: 'dropdown',
    options: ['English','Mandarin Chinese','Hindi','Spanish','French','Arabic','Bengali','Russian','Portuguese','German','Other'],
  },
  {
    variable: 'var_industry',
    question: 'What do you do for a living?',
    type: 'dropdown',
    options: ['Law & Legal Services','Business & Finance','Tech & IT','Marketing & Media','Education & Research','Healthcare & Sciences','Student','Other'],
  },
];

// --- POST-GAME INTERVIEW ---
const POST_SURVEY = [
  { variable: 'TI-WTD_1', question: 'When an important legal issue or problem arises, I would feel comfortable depending on the information provided by the AI assistant.', type: 'likert' },
  { variable: 'TI-WTD_2', question: 'I can always rely on the AI assistant in a tough legal situation.', type: 'likert' },
  { variable: 'TI-WTD_3', question: 'I feel that I could count on the AI assistant to help with a crucial legal problem.', type: 'likert' },
  { variable: 'TI-WTD_4', question: 'Faced with a difficult legal question that required me to hire a lawyer, I would use the AI assistant.', type: 'likert' },
  { variable: 'TI-WTD_5', question: 'If I had a challenging legal problem, I would want to use the AI assistant again.', type: 'likert' },
  { variable: 'Ex-Q_1',   question: 'If I had to perform a similar task in my professional life, I would choose to use this AI assistant again.', type: 'likert' },
  { variable: 'Ex-Q_2',   question: 'I would recommend this specific AI interface to my colleagues for consultancy work.', type: 'likert' },
];

// --- CATEGORY → DOMAIN DISPLAY LABEL ---
const CATEGORY_LABELS = {
  'Consumer Protection': 'Consumer Protection',
  'Workforce Regulation': 'Workforce Regulation',
  'Company Taxation':    'Company Taxation',
};

// ═══════════════════════════════════════════════════════════════
// 2. GAME STATE
// ═══════════════════════════════════════════════════════════════

const state = {
  group:          'A',      // A, B, C, D
  money:          CONFIG.startingMoney,
  hearts:         CONFIG.startingHearts,
  currentDay:     1,
  currentCaseIdx: 0,        // index into MASTER_SCRIPT
  timerVal:       CONFIG.timeLimit,
  timerInterval:  null,
  caseStartTime:  null,
  sourceClicked:  false,
  handbookReturnScreen: null,
  correctCount:   0,

  // Per-case log entries
  log: [],

  // Survey answers
  preSurveyAnswers:  {},
  postSurveyAnswers: {},

  // Day summary buffer
  dayResults: [],
};

// ═══════════════════════════════════════════════════════════════
// 3. UTILITY HELPERS
// ═══════════════════════════════════════════════════════════════

function $(id) { return document.getElementById(id); }

function showScreen(id) {
  document.querySelectorAll('.screen').forEach(s => {
    s.classList.remove('active');
    s.style.display = 'none';
  });
  const el = $(id);
  el.style.display = 'flex';
  el.classList.add('active');
  el.scrollTop = 0;
}
// ═══════════════════════════════════════════════════════════════
// SCENE CONTROLLER — pixel-art SVG office animations
// ═══════════════════════════════════════════════════════════════

window.sceneController = (function() {
  let _state = 'hidden';
  let _bossX  = 320; // starts off-screen right
  let _walkRAF = null;

  // Target X positions (boss-group SVG translate X)
  const POS = {
    hidden:   350,  // fully off-screen right
    peeking:  292,  // just peeking from right edge
    walking:  292,  // start of walk (same as peeking)
    looming:  195,  // standing right behind avatar
    outburst: 195,  // same position, fire + shake
  };

  function getBossGroup() { return document.getElementById('boss-group'); }
  function getOutfire1()  { return document.getElementById('outburst-fire'); }
  function getOutfire2()  { return document.getElementById('outburst-fire2'); }
  function getFootTap()   { return document.getElementById('boss-foot-tap'); }
  function getBrowL()     { return document.getElementById('boss-brow-l'); }
  function getBrowR()     { return document.getElementById('boss-brow-r'); }
  function getBossMouth() { return document.getElementById('boss-mouth'); }

  function setTranslate(el, x, y) {
    if (!el) return;
    el.setAttribute('transform', `translate(${Math.round(x)}, ${y})`);
  }

  function cancelWalk() {
    if (_walkRAF) { cancelAnimationFrame(_walkRAF); _walkRAF = null; }
  }

  function walkTo(targetX, bossY, speed, onDone) {
    cancelWalk();
    function step() {
      const dx = targetX - _bossX;
      if (Math.abs(dx) < 1) {
        _bossX = targetX;
        setTranslate(getBossGroup(), _bossX, bossY);
        if (onDone) onDone();
        return;
      }
      _bossX += dx * speed;
      setTranslate(getBossGroup(), _bossX, bossY);
      _walkRAF = requestAnimationFrame(step);
    }
    _walkRAF = requestAnimationFrame(step);
  }

  function resetFaceNeutral() {
    const bl = getBrowL(), br = getBrowR(), m = getBossMouth();
    if (bl) bl.setAttribute('y', '5');
    if (br) br.setAttribute('y', '5');
    if (m)  { m.setAttribute('width', '6'); m.setAttribute('fill', '#a06050'); }
  }

  function setFaceAngry() {
    const bl = getBrowL(), br = getBrowR(), m = getBossMouth();
    // Angled angry brows
    if (bl) bl.setAttribute('y', '6');
    if (br) br.setAttribute('y', '4');
    if (m)  { m.setAttribute('width', '4'); m.setAttribute('fill', '#882020'); }
  }

  function showFire(on) {
    const f1 = getOutfire1(), f2 = getOutfire2();
    if (f1) f1.setAttribute('opacity', on ? '1' : '0');
    if (f2) f2.setAttribute('opacity', on ? '1' : '0');
  }

  function showFootTap(on) {
    const ft = getFootTap();
    if (ft) ft.setAttribute('opacity', on ? '1' : '0');
  }

  function setState(s) {
    if (_state === s) return;
    _state = s;

    switch(s) {

      case 'hidden':
        cancelWalk();
        _bossX = POS.hidden;
        setTranslate(getBossGroup(), _bossX, 88);
        showFire(false);
        showFootTap(false);
        resetFaceNeutral();
        break;

      case 'peeking':
        // Slide boss in from right so only head+shoulder peek
        cancelWalk();
        _bossX = POS.peeking;
        setTranslate(getBossGroup(), _bossX, 88);
        showFire(false);
        showFootTap(false);
        resetFaceNeutral();
        break;

      case 'walking':
        // Animate walk from peek position to behind avatar
        showFire(false);
        showFootTap(false);
        resetFaceNeutral();
        // Start from peeking position if not already there
        if (_bossX > POS.peeking + 5) {
          _bossX = POS.peeking;
          setTranslate(getBossGroup(), _bossX, 88);
        }
        walkTo(POS.looming, 88, 0.06, () => {
          // Arrived — now loom
          showFootTap(true);
          setFaceAngry();
        });
        break;

      case 'looming':
        cancelWalk();
        _bossX = POS.looming;
        setTranslate(getBossGroup(), _bossX, 88);
        showFire(false);
        showFootTap(true);
        setFaceAngry();
        break;

      case 'outburst':
        cancelWalk();
        _bossX = POS.looming;
        setTranslate(getBossGroup(), _bossX, 88);
        showFire(true);
        showFootTap(false);
        setFaceAngry();
        // Auto-reset after 2.2s
        setTimeout(() => {
          showFire(false);
          resetFaceNeutral();
        }, 2200);
        break;
    }
  }

  // Update timer display colour in SVG
  function updateTimerSvg(val) {
    const el = document.getElementById('timer-display');
    if (!el) return;
    el.textContent = val;
    // Colour: green→amber→red
    if (val <= 20)      el.setAttribute('fill', '#ff2200');
    else if (val <= 40) el.setAttribute('fill', '#ff9900');
    else                el.setAttribute('fill', '#cc2222');

    // Timer wrap pulse class
    const wrap = document.getElementById('pixel-timer');
    if (!wrap) return;
    wrap.classList.remove('warn-amber', 'warn-red');
    if (val <= 20)      wrap.classList.add('warn-red');
    else if (val <= 40) wrap.classList.add('warn-amber');
  }

  return { setState, updateTimerSvg };
})();



function getUrlGroup() {
  const params = new URLSearchParams(window.location.search);
  const g = (params.get('group') || 'A').toUpperCase();
  return ['A','B','C','D'].includes(g) ? g : 'A';
}

function heartsDisplay(n) {
  if (n > 0) return '❤️'.repeat(n);
  if (n === 0) return '0';
  return '💔 ' + n; // negatives shown as 💔 -2 etc
}

function updateHUD() {
  $('hud-day').textContent    = state.currentDay;
  $('hud-money').textContent  = state.money.toLocaleString();
  $('hud-hearts').textContent = heartsDisplay(state.hearts);
  // case display: count regular cases in day
  const dayCases = MASTER_SCRIPT.filter(c => c.day === state.currentDay && !c.isAttentionCheck);
  const casesDone = state.dayResults.filter(r => !r.isAttentionCheck).length;
  $('hud-case').textContent   = `${casesDone + 1}/${dayCases.length}`;
}

// ═══════════════════════════════════════════════════════════════
// 4. SURVEY BUILDERS
// ═══════════════════════════════════════════════════════════════

function buildPreSurvey() {
  const container = $('pre-survey-form');
  container.innerHTML = '';

  PRE_SURVEY.forEach((q, i) => {
    const wrap = document.createElement('div');
    wrap.className = 'survey-question';

    const label = document.createElement('label');
    label.innerHTML = `<strong>${i + 1}. ${q.question}</strong>`;

    let input;
    if (q.type === 'dropdown') {
      input = document.createElement('select');
      input.id = `pre_${q.variable}`;
      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = '— Select an option —';
      placeholder.disabled = true;
      placeholder.selected = true;
      input.appendChild(placeholder);
      q.options.forEach(opt => {
        const o = document.createElement('option');
        o.value = opt;
        o.textContent = opt;
        input.appendChild(o);
      });
      input.addEventListener('change', () => {
        state.preSurveyAnswers[q.variable] = input.value;
      });
    }

    wrap.appendChild(label);
    wrap.appendChild(input);
    container.appendChild(wrap);
  });
}

function buildPostSurvey() {
  const container = $('post-survey-form');
  container.innerHTML = '';

  POST_SURVEY.forEach((q, i) => {
    const wrap = document.createElement('div');
    wrap.className = 'survey-question';

    const label = document.createElement('label');
    label.innerHTML = `<strong>${i + 1}. ${q.question}</strong>`;
    wrap.appendChild(label);

    if (q.type === 'likert') {
      const row = document.createElement('div');
      row.className = 'likert-row';
      for (let v = 1; v <= 7; v++) {
        const btn = document.createElement('button');
        btn.className = 'likert-btn';
        btn.textContent = v;
        btn.dataset.val = v;
        btn.addEventListener('click', () => {
          row.querySelectorAll('.likert-btn').forEach(b => b.classList.remove('selected'));
          btn.classList.add('selected');
          state.postSurveyAnswers[q.variable] = v;
        });
        row.appendChild(btn);
      }
      const labels = document.createElement('div');
      labels.className = 'likert-labels';
      labels.innerHTML = '<span>Strongly Disagree</span><span>Strongly Agree</span>';
      wrap.appendChild(row);
      wrap.appendChild(labels);
    }

    container.appendChild(wrap);
  });
}

// ═══════════════════════════════════════════════════════════════
// 5. HANDBOOK
// ═══════════════════════════════════════════════════════════════

function buildHandbook(highlightId) {
  const tabs = $('handbook-tabs');
  const articles = $('handbook-articles');
  tabs.innerHTML = '';
  articles.innerHTML = '';

  const categories = [...new Set(HANDBOOK.map(a => a.category))];
  let activeTab = categories[0];

  // If highlight specified, open its category
  if (highlightId) {
    const target = HANDBOOK.find(a => a.id === highlightId);
    if (target) activeTab = target.category;
  }

  categories.forEach(cat => {
    const tab = document.createElement('div');
    tab.className = 'handbook-tab' + (cat === activeTab ? ' active' : '');
    tab.textContent = cat;
    tab.addEventListener('click', () => {
      tabs.querySelectorAll('.handbook-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      showCategory(cat);
    });
    tabs.appendChild(tab);
  });

  HANDBOOK.forEach(article => {
    const div = document.createElement('div');
    div.className = 'handbook-article' + (article.category === activeTab ? ' visible' : '');
    if (highlightId && article.id === highlightId) div.classList.add('highlighted');
    div.id = `article-${article.id}`;
    div.innerHTML = `
      <div class="article-id">${article.id}</div>
      <div class="article-cat">${article.category}</div>
      <div class="article-text">${article.text}</div>
    `;
    articles.appendChild(div);
  });

  // Scroll to highlighted article
  if (highlightId) {
    requestAnimationFrame(() => {
      const el = $(`article-${highlightId}`);
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
  }
}

function showCategory(cat) {
  document.querySelectorAll('.handbook-article').forEach(el => {
    el.classList.toggle('visible', el.querySelector('.article-cat').textContent === cat);
  });
}

// ═══════════════════════════════════════════════════════════════
// 6. CASE RENDERING
// ═══════════════════════════════════════════════════════════════

function getHandbookArticle(id) {
  return HANDBOOK.find(a => a.id === id) || null;
}

function getAIText(caseData) {
  const isAssertive = (state.group === 'A' || state.group === 'B');
  return isAssertive ? caseData.aiAssertive : caseData.aiHedged;
}

function renderCase() {
  const caseData = MASTER_SCRIPT[state.currentCaseIdx];
  if (!caseData) { endGame(); return; }

  state.sourceClicked = false;
  state.caseStartTime = Date.now();

  // Question
  $('case-question').textContent = caseData.question;

  // AI advice
  const aiText = getAIText(caseData);
  $('ai-advice').textContent = aiText;

  // Source button (Groups B and D, non-attention-check, has handbookRef)
  const showSource = (state.group === 'B' || state.group === 'D') &&
                     !caseData.isAttentionCheck &&
                     caseData.handbookRef;
  const srcBtn = $('btn-source');
  if (showSource) {
    srcBtn.style.display = 'inline-block';
    srcBtn.classList.remove('clicked');
    srcBtn.textContent = '📖 CHECK SOURCE';
    srcBtn.disabled = false;
  } else {
    srcBtn.style.display = 'none';
  }

  // Options
  const grid = $('options-grid');
  grid.innerHTML = '';
  ['A','B','C','D'].forEach(letter => {
    const btn = document.createElement('button');
    btn.className = 'option-btn';
    btn.dataset.label = `OPTION ${letter}`;
    btn.textContent = caseData.options[letter];
    btn.addEventListener('click', () => submitAnswer(letter));
    grid.appendChild(btn);
  });

  updateHUD();
  resetTimer();
  resetManagerAvatar();
  showScreen('screen-game');
}

// ═══════════════════════════════════════════════════════════════
// 7. TIMER
// ═══════════════════════════════════════════════════════════════

function resetTimer() {
  clearInterval(state.timerInterval);
  state.timerVal = CONFIG.timeLimit;
  updateTimerDisplay();
  setBossState('hidden');
  state.timerInterval = setInterval(tickTimer, 1000);
}

function updateTimerDisplay() {
  const t = state.timerVal;
  // Update SVG timer (new scene)
  if (window.sceneController) window.sceneController.updateTimerSvg(t);
}

function tickTimer() {
  state.timerVal--;
  updateTimerDisplay();

  // Boss progression: clock counts 60→0
  // 40s left (20s elapsed)  → Boss peeks from corner
  // 20s left (40s elapsed)  → Boss walks to desk, stands impatiently
  if (state.timerVal === 40) setBossState('peeking');
  if (state.timerVal === 20) setBossState('walking');

  if (state.timerVal <= 0) {
    clearInterval(state.timerInterval);
    timeOut();
  }
}

function stopTimer() {
  clearInterval(state.timerInterval);
}

// ═══════════════════════════════════════════════════════════════
// 8. SCENE / BOSS ANIMATION CONTROLLER
// ═══════════════════════════════════════════════════════════════

// Boss states: 'hidden' | 'peeking' | 'walking' | 'looming' | 'outburst'
function setBossState(s) {
  const scene = $('office-scene');
  if (!scene) return;
  scene.dataset.bossState = s;
  // Delegate visual updates to the SVG scene's own logic
  if (window.sceneController) window.sceneController.setState(s);
}

function resetManagerAvatar() {
  setBossState('hidden');
}

function triggerOutburst() {
  setBossState('outburst');
  const el = $('screen-game');
  el.classList.add('shaking');
  setTimeout(() => el.classList.remove('shaking'), 800);
  // Auto-clear outburst after 2s
  setTimeout(() => setBossState('hidden'), 2000);
}

function triggerScreenShake() {
  const el = $('screen-game');
  el.classList.add('shaking');
  setTimeout(() => el.classList.remove('shaking'), 700);
}

// ═══════════════════════════════════════════════════════════════
// 9. ANSWER HANDLING
// ═══════════════════════════════════════════════════════════════

function submitAnswer(chosen) {
  stopTimer();
  const caseData = MASTER_SCRIPT[state.currentCaseIdx];
  const deliberationTime = Math.round((Date.now() - state.caseStartTime) / 1000);
  const correct = chosen === caseData.correctAnswer;

  // Disable all buttons to prevent double-click
  document.querySelectorAll('.option-btn').forEach(b => b.disabled = true);

  // Log entry
  const logEntry = {
    Session_Time:      new Date().toISOString(),
    Group:             state.group,
    Case_ID:           caseData.id,
    Day:               caseData.day,
    Is_Attention_Check:caseData.isAttentionCheck,
    AI_Style:          (state.group === 'A' || state.group === 'B') ? 'Assertive' : 'Hedged',
    Source_Available:  (state.group === 'B' || state.group === 'D') && !caseData.isAttentionCheck && !!caseData.handbookRef,
    Source_Checked:    state.sourceClicked,
    Chosen_Option:     chosen,
    Correct_Answer:    caseData.correctAnswer,
    Is_Correct:        correct,
    Deliberation_Time: deliberationTime,
    Money_Before:      state.money,
    Hearts_Before:     state.hearts,
  };

  if (caseData.isAttentionCheck) {
    // Attention check — record but don't penalize
    logEntry.Money_After  = state.money;
    logEntry.Hearts_After = state.hearts;
    state.log.push(logEntry);
    state.dayResults.push({ caseData, chosen, correct, isAttentionCheck: true, deliberationTime });
    showAttentionResult(correct, chosen, caseData, () => advanceCase());
    return;
  }

  // Apply score changes
  if (correct) {
    state.money  += CONFIG.rewardMoney;
    state.hearts  = Math.min(5, state.hearts + CONFIG.rewardHearts);
    state.correctCount++;
  } else {
    state.money  -= CONFIG.penaltyMoney;
    state.hearts  = Math.max(0, state.hearts - CONFIG.penaltyHearts);
  }
  logEntry.Money_After  = state.money;
  logEntry.Hearts_After = state.hearts;
  state.log.push(logEntry);
  state.dayResults.push({ caseData, chosen, correct, isAttentionCheck: false, deliberationTime });

  showResultOverlay(correct, chosen, caseData, false, () => advanceCase());
}

function timeOut() {
  const caseData = MASTER_SCRIPT[state.currentCaseIdx];
  const deliberationTime = CONFIG.timeLimit;

  // Timeout counts as wrong
  state.money  -= CONFIG.penaltyMoney;
  state.hearts  = Math.max(0, state.hearts - CONFIG.penaltyHearts);

  const logEntry = {
    Session_Time:      new Date().toISOString(),
    Group:             state.group,
    Case_ID:           caseData.id,
    Day:               caseData.day,
    Is_Attention_Check:caseData.isAttentionCheck,
    AI_Style:          (state.group === 'A' || state.group === 'B') ? 'Assertive' : 'Hedged',
    Source_Available:  (state.group === 'B' || state.group === 'D') && !caseData.isAttentionCheck && !!caseData.handbookRef,
    Source_Checked:    state.sourceClicked,
    Chosen_Option:     'TIMEOUT',
    Correct_Answer:    caseData.correctAnswer,
    Is_Correct:        false,
    Deliberation_Time: deliberationTime,
    Money_Before:      state.money + CONFIG.penaltyMoney,
    Hearts_Before:     state.hearts + CONFIG.penaltyHearts,
    Money_After:       state.money,
    Hearts_After:      state.hearts,
  };
  state.log.push(logEntry);
  if (!caseData.isAttentionCheck) {
    state.dayResults.push({ caseData, chosen: 'TIMEOUT', correct: false, isAttentionCheck: false, deliberationTime });
  }

  document.querySelectorAll('.option-btn').forEach(b => b.disabled = true);
  triggerOutburst();
  showResultOverlay(false, 'TIMEOUT', caseData, true, () => advanceCase());
}

// ═══════════════════════════════════════════════════════════════
// 10. RESULT OVERLAYS
// ═══════════════════════════════════════════════════════════════

function showResultOverlay(correct, chosen, caseData, isTimeout, onDone) {
  const overlay  = $('result-overlay');
  const icon     = $('result-icon');
  const titleEl  = $('result-title');
  const detail   = $('result-detail');
  const delta    = $('result-delta');

  overlay.style.display = 'flex';

  if (isTimeout) {
    icon.textContent = '⏰';
    titleEl.textContent = 'TIME\'S UP';
    titleEl.className = 'result-title timeout';
    detail.textContent = `Correct answer: ${caseData.correctAnswer}`;
    delta.textContent  = `-${CONFIG.penaltyMoney}€ · -${CONFIG.penaltyHearts}❤️`;
    delta.className    = 'result-delta negative';
  } else if (correct) {
    icon.textContent = '✅';
    titleEl.textContent = 'CORRECT!';
    titleEl.className = 'result-title correct';
    detail.textContent = `+${CONFIG.rewardMoney}€ earned`;
    delta.textContent  = `+${CONFIG.rewardMoney}€ · +${CONFIG.rewardHearts}❤️`;
    delta.className    = 'result-delta positive';
  } else {
    icon.textContent = '❌';
    titleEl.textContent = 'WRONG';
    titleEl.className = 'result-title wrong';
    detail.textContent = `Correct: ${caseData.correctAnswer} — ${caseData.options[caseData.correctAnswer]}`;
    delta.textContent  = `-${CONFIG.penaltyMoney}€ · -${CONFIG.penaltyHearts}❤️`;
    delta.className    = 'result-delta negative';
    triggerOutburst();
  }

  setTimeout(() => {
    overlay.style.display = 'none';
    onDone();
  }, 2200);
}

function showAttentionResult(correct, chosen, caseData, onDone) {
  const overlay = $('attention-overlay');
  const detail  = $('attention-detail');
  const res     = $('attention-result');

  overlay.style.display = 'flex';
  detail.textContent = 'This was an attention check — no points are deducted.';
  res.textContent    = correct ? '✅ Correct answer selected.' : `❌ Correct answer was: ${caseData.correctAnswer} — ${caseData.options[caseData.correctAnswer]}`;
  res.className      = 'result-delta ' + (correct ? 'positive' : 'negative');

  setTimeout(() => {
    overlay.style.display = 'none';
    onDone();
  }, 2400);
}

// ═══════════════════════════════════════════════════════════════
// 11. CASE FLOW / DAY FLOW
// ═══════════════════════════════════════════════════════════════

function advanceCase() {
  state.currentCaseIdx++;

  // Are there more cases in the current day?
  const nextCase = MASTER_SCRIPT[state.currentCaseIdx];
  if (!nextCase) {
    // All days done
    showEndOfDay(true);
    return;
  }

  if (nextCase.day > state.currentDay) {
    // New day starts
    showEndOfDay(false);
  } else {
    renderCase();
  }
}

function showEndOfDay(isLastDay) {
  const day = state.currentDay;
  const endMsgs = {
    1: 'Shift over. Let\'s review your daily performance.',
    2: 'Shift over. Let\'s review your daily performance.',
    3: 'Shift over. Let\'s review your daily performance.',
    4: 'Shift over. Let\'s review your daily performance.',
    5: 'Protocol Completed. Calculating final score…',
  };

  $('day-end-title').textContent = `END OF DAY ${day}`;
  $('day-end-msg').textContent   = endMsgs[day] || endMsgs[1];
  $('end-money').textContent     = state.money.toLocaleString();
  $('end-hearts').textContent    = heartsDisplay(state.hearts);

  // Render day results
  const resultsEl = $('day-end-results');
  resultsEl.innerHTML = '';
  state.dayResults.forEach(r => {
    const row = document.createElement('div');
    const cls  = r.isAttentionCheck ? 'dr-attention' : (r.chosen === 'TIMEOUT' ? 'dr-timeout' : (r.correct ? 'dr-correct' : 'dr-wrong'));
    row.className = `day-result-row ${cls}`;
    const icon   = r.isAttentionCheck ? '🔔' : (r.chosen === 'TIMEOUT' ? '⏰' : (r.correct ? '✅' : '❌'));
    const label  = r.isAttentionCheck ? 'ATTENTION' : (r.chosen === 'TIMEOUT' ? 'TIMEOUT' : (r.correct ? 'CORRECT' : 'WRONG'));
    const qShort = r.caseData.question.slice(0, 55) + (r.caseData.question.length > 55 ? '…' : '');
    row.innerHTML = `
      <span class="dr-icon">${icon}</span>
      <span class="dr-q">${qShort}</span>
      <span class="dr-result">${label}</span>
    `;
    resultsEl.appendChild(row);
  });

  // Reset day results buffer for next day
  state.dayResults = [];

  const btnNext = $('btn-next-day');
  if (isLastDay) {
    btnNext.textContent = 'Continue ▶';
    btnNext.onclick = () => showScreen('screen-post-survey');
  } else {
    btnNext.textContent = 'Next Day ▶';
    btnNext.onclick = () => {
      state.currentDay++;
      showDayIntro(state.currentDay);
    };
  }
  showScreen('screen-day-end');
}

function showDayIntro(day) {
  const msgs = {
    1: 'Day 1: Get ready.',
    2: 'Day 2: Stay sharp.',
    3: 'Day 3: The boss is watching.',
    4: 'Day 4: Don\'t let anything slip.',
    5: 'Day 5: Make it count!',
  };
  $('day-intro-badge').textContent  = `DAY ${day}`;
  $('day-intro-msg').textContent    = msgs[day] || `Day ${day}`;
  const dayCases = MASTER_SCRIPT.filter(c => c.day === day && !c.isAttentionCheck);
  const first = MASTER_SCRIPT.findIndex(c => c.day === day) + 1;
  $('day-case-range').textContent   = `${first}–${first + dayCases.length - 1}`;
  $('intro-money').textContent      = state.money.toLocaleString();
  $('intro-hearts').textContent     = state.hearts;
  showScreen('screen-day-intro');
}

function endGame() {
  showScreen('screen-post-survey');
}

// ═══════════════════════════════════════════════════════════════
// 12. SOURCE BUTTON HANDLER
// ═══════════════════════════════════════════════════════════════

const app = {
  openHandbookFree() {
    // Free handbook access — timer keeps running, no source_checked flag
    state.handbookReturnScreen = 'screen-game';
    buildHandbook(null);
    showScreen('screen-handbook');
  },

  handleSourceClick() {
    if (state.sourceClicked) return; // idempotent
    state.sourceClicked = true;
    const srcBtn = $('btn-source');
    srcBtn.classList.add('clicked');
    srcBtn.textContent = '✅ SOURCE CHECKED';
    srcBtn.disabled = true;

    const caseData = MASTER_SCRIPT[state.currentCaseIdx];
    state.handbookReturnScreen = 'screen-game';
    stopTimer(); // Pause while viewing handbook
    buildHandbook(caseData.handbookRef);
    showScreen('screen-handbook');
  }
};

// ═══════════════════════════════════════════════════════════════
// 13. DEBRIEF / DATA DOWNLOAD
// ═══════════════════════════════════════════════════════════════

function showDebrief() {
  $('final-money').textContent        = state.money.toLocaleString();
  $('final-hearts-display').textContent = heartsDisplay(state.hearts);
  $('final-correct').textContent      = `${state.correctCount}/15`;
  showScreen('screen-debrief');
}

// ═══════════════════════════════════════════════════════════════
// 14. INIT & EVENT WIRING
// ═══════════════════════════════════════════════════════════════

function init() {
  state.group = getUrlGroup();

  // Build surveys
  buildPreSurvey();
  buildPostSurvey();

  // Build handbook tabs (initially without highlight)
  buildHandbook(null);

  // --- Disclaimer ---
  $('btn-agree').addEventListener('click', () => {
    showScreen('screen-pre-survey');
  });

  // --- Pre-survey submit ---
  $('btn-pre-submit').addEventListener('click', () => {
    // Validate all dropdowns selected
    let allFilled = true;
    PRE_SURVEY.forEach(q => {
      const sel = $(`pre_${q.variable}`);
      if (!sel || !sel.value) allFilled = false;
    });
    if (!allFilled) {
      alert('Please answer all questions before continuing.');
      return;
    }
    // Read final values
    PRE_SURVEY.forEach(q => {
      const sel = $(`pre_${q.variable}`);
      if (sel) state.preSurveyAnswers[q.variable] = sel.value;
    });
    showDayIntro(1);
  });

  // --- Day intro: begin shift ---
  $('btn-begin-day').addEventListener('click', () => {
    renderCase();
  });

  // --- Handbook back button ---
  $('btn-close-handbook').addEventListener('click', () => {
    const ret = state.handbookReturnScreen;
    if (ret === 'screen-game') {
      showScreen('screen-game');
      // Resume timer from where it paused
      state.timerInterval = setInterval(tickTimer, 1000);
    } else {
      showScreen(ret || 'screen-game');
    }
    state.handbookReturnScreen = null;
  });

  // --- Post-survey submit ---
  $('btn-post-submit').addEventListener('click', () => {
    const unanswered = POST_SURVEY.filter(q => state.postSurveyAnswers[q.variable] === undefined);
    if (unanswered.length > 0) {
      alert('Please answer all questions before continuing.');
      return;
    }
    showDebrief();
  });

  // Expose app globally (for inline onclick)
  window.app = app;

  showScreen('screen-disclaimer');
}

document.addEventListener('DOMContentLoaded', init);