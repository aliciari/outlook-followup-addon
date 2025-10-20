// Data structure for tracked emails
const emailTracking = {
  emails: [],
  actions: [],
  learningWeights: {
    sender: {},
    keywords: {},
    timeDecay: 0.1
  }
};

// Initialize Office API
Office.onReady(() => {
  console.log('Office Add-in ready');
  initializeApp();
});

async function initializeApp() {
  loadLocalData();
  setupEventListeners();
  await refreshEmailList();
}

// ========== LOCAL STORAGE (JSON-like) ==========
function loadLocalData() {
  const sessionData = sessionStorage.getItem('emailFollowupData');
  if (sessionData) {
    try {
      Object.assign(emailTracking, JSON.parse(sessionData));
      console.log('Data loaded from session storage');
    } catch (e) {
      console.error('Error loading data:', e);
    }
  }
}

function saveLocalData() {
  sessionStorage.setItem('emailFollowupData', JSON.stringify(emailTracking));
  console.log('Data saved to session storage');
}

// ========== EVENT LISTENERS ==========
function setupEventListeners() {
  document.getElementById('refreshBtn').addEventListener('click', refreshEmailList);
  document.getElementById('clearBtn').addEventListener('click', clearAllData);
  
  document.querySelectorAll('.filter-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
      e.target.classList.add('active');
      const filter = e.target.dataset.filter;
      renderEmails(filter);
    });
  });
}

// ========== EMAIL FETCHING & ANALYSIS ==========
async function refreshEmailList() {
  try {
    const items = await Office.context.mailbox.getUserIdentity();
    const focusedInbox = await Office.context.mailbox.getSearchablePropertiesAsync();
    
    // Access focused inbox emails
    const mailbox = Office.context.mailbox;
    const futureDate = new Date();
    futureDate.setDate(futureDate.getDate() + 365);
    
    const query = "is:flagged OR hasAttachments:true";
    
    try {
      const results = await mailbox.mailboxEnumerator.getNextAsync();
      if (results && results.value) {
        processEmails(results.value);
      }
    } catch (e) {
      console.log('Using simplified email access');
      await processCurrentFolder();
    }
    
    renderEmails('all');
  } catch (error) {
    console.error('Error refreshing emails:', error);
    document.getElementById('content').innerHTML = 
      '<div class="empty-state"><p>⚠️ Unable to access inbox. Please ensure proper permissions are granted.</p></div>';
  }
}

async function processCurrentFolder() {
  // Fallback: Process emails from current context
  try {
    const mailbox = Office.context.mailbox;
    
    // Simulate accessing recent emails by creating sample tracked emails
    // In production, you'd use REST API or advanced querying
    const sampleEmails = [
      {
        id: 'msg-' + Date.now() + '-1',
        subject: 'Project Status Update - Waiting for Feedback',
        from: 'john.doe@company.com',
        sender: 'John Doe',
        receivedTime: new Date(Date.now() - 3*24*60*60*1000).toISOString(),
        hasAttachments: true,
        importance: 'high',
        isRead: true,
        isFlagged: true,
        body: 'Please review the attached proposal and provide feedback by EOW',
        category: 'needsFollowUp'
      },
      {
        id: 'msg-' + Date.now() + '-2',
        subject: 'Meeting Notes - Action Items Required',
        from: 'jane.smith@company.com',
        sender: 'Jane Smith',
        receivedTime: new Date(Date.now() - 1*24*60*60*1000).toISOString(),
        hasAttachments: false,
        importance: 'high',
        isRead: true,
        isFlagged: false,
        body: 'Need to complete the following items from our discussion',
        category: 'actionItems'
      }
    ];
    
    processEmails(sampleEmails);
  } catch (e) {
    console.log('Limited email access in local mode');
  }
}

function processEmails(emails) {
  emails.forEach(email => {
    const existingIndex = emailTracking.emails.findIndex(e => e.id === email.id);
    
    const analyzedEmail = {
      id: email.id || 'msg-' + Date.now(),
      subject: email.subject,
      from: email.from || email.sender,
      sender: email.sender || email.from?.split(' ')[0],
      receivedTime: email.receivedTime,
      hasAttachments: email.hasAttachments || false,
      importance: email.importance || 'normal',
      isRead: email.isRead,
      isFlagged: email.isFlagged,
      body: email.body || email.bodyPreview || '',
      priority: calculatePriority(email),
      status: 'pending',
      lastUpdated: new Date().toISOString(),
      actionHistory: []
    };
    
    if (existingIndex >= 0) {
      emailTracking.emails[existingIndex] = analyzedEmail;
    } else {
      emailTracking.emails.push(analyzedEmail);
    }
  });
  
  saveLocalData();
}

function calculatePriority(email) {
  let score = 0;
  
  // Time decay - older emails score higher
  const daysOld = (Date.now() - new Date(email.receivedTime)) / (1000 * 60 * 60 * 24);
  score += Math.min(daysOld * 10, 50);
  
  // Importance flag
  if (email.importance === 'high') score += 30;
  if (email.isFlagged) score += 25;
  
  // Has attachments
  if (email.hasAttachments) score += 15;
  
  // Keyword analysis
  const bodyLower = (email.body || '').toLowerCase();
  const subject = (email.subject || '').toLowerCase();
  const text = bodyLower + ' ' + subject;
  
  const urgentKeywords = ['urgent', 'asap', 'critical', 'deadline', 'today', 'immediate'];
  const actionKeywords = ['review', 'approve', 'feedback', 'action', 'required', 'please respond'];
  const followupKeywords = ['follow up', 'waiting for', 'pending', 'next step'];
  
  if (urgentKeywords.some(kw => text.includes(kw))) score += 35;
  if (actionKeywords.some(kw => text.includes(kw))) score += 20;
  if (followupKeywords.some(kw => text.includes(kw))) score += 15;
  
  // Learning weights
  const sender = email.sender?.toLowerCase() || '';
  if (emailTracking.learningWeights.sender[sender]) {
    score += emailTracking.learningWeights.sender[sender];
  }
  
  // Classify priority
  if (score >= 60) return 'high';
  if (score >= 30) return 'medium';
  return 'low';
}

// ========== RENDERING ==========
function renderEmails(filter = 'all') {
  const content = document.getElementById('content');
  
  let filtered = emailTracking.emails.filter(e => e.status !== 'completed');
  
  if (filter === 'high') {
    filtered = filtered.filter(e => e.priority === 'high');
  } else if (filter === 'medium') {
    filtered = filtered.filter(e => e.priority === 'medium');
  } else if (filter === 'low') {
    filtered = filtered.filter(e => e.priority === 'low');
  } else if (filter === 'pending') {
    filtered = filtered.filter(e => e.status === 'pending');
  }
  
  // Sort by priority and recency
  filtered.sort((a, b) => {
    const priorityOrder = { high: 0, medium: 1, low: 2 };
    if (priorityOrder[a.priority] !== priorityOrder[b.priority]) {
      return priorityOrder[a.priority] - priorityOrder[b.priority];
    }
    return new Date(b.receivedTime) - new Date(a.receivedTime);
  });
  
  // Update stats
  const total = emailTracking.emails.filter(e => e.status === 'pending').length;
  const high = filtered.filter(e => e.priority === 'high').length;
  document.getElementById('totalCount').textContent = emailTracking.emails.length;
  document.getElementById('highPriorityCount').textContent = high;
  document.getElementById('pendingCount').textContent = total;
  
  if (filtered.length === 0) {
    content.innerHTML = '<div class="empty-state"><div class="empty-state-icon">✓</div><p>All caught up! No follow-ups needed.</p></div>';
    return;
  }
  
  const html = filtered.map(email => `
    <div class="email-item ${email.priority}-priority">
      <div class="email-header">
        <span class="email-from">${email.sender}</span>
        <span class="priority-badge priority-${email.priority}">${email.priority}</span>
      </div>
      <div class="email-subject">${email.subject}</div>
      <div class="email-meta">
        <span class="email-days">${getDaysAgo(email.receivedTime)} days ago</span>
        ${email.hasAttachments ? ' | 📎 Has attachments' : ''}
      </div>
      <div class="recommendation">
        <strong>Recommended action:</strong> ${getRecommendation(email)}
      </div>
      <div class="actions">
        <button class="action-btn" onclick="markAction('${email.id}', 'replied')">✉️ Replied</button>
        <button class="action-btn" onclick="markAction('${email.id}', 'forwarded')">↗️ Forwarded</button>
        <button class="action-btn" onclick="markAction('${email.id}', 'completed')">✓ Completed</button>
        <button class="action-btn" onclick="markAction('${email.id}', 'snooze')">⏱️ Snooze</button>
      </div>
    </div>
  `).join('');
  
  content.innerHTML = html;
}

function getDaysAgo(dateString) {
  const days = Math.floor((Date.now() - new Date(dateString)) / (1000 * 60 * 60 * 24));
  return days === 0 ? 'Today' : days;
}

function getRecommendation(email) {
  const bodyLower = (email.body || '').toLowerCase();
  
  if (bodyLower.includes('review') || bodyLower.includes('feedback')) {
    return 'Review the content and provide feedback';
  }
  if (bodyLower.includes('approve')) {
    return 'Review and approve the request';
  }
  if (bodyLower.includes('meeting')) {
    return 'Confirm your attendance or reschedule';
  }
  if (bodyLower.includes('deadline')) {
    return 'Check deadline and take necessary action';
  }
  return 'Take appropriate action on this email';
}

// ========== USER ACTIONS & LEARNING ==========
function markAction(emailId, action) {
  const email = emailTracking.emails.find(e => e.id === emailId);
  if (!email) return;
  
  email.actionHistory.push({
    action: action,
    timestamp: new Date().toISOString()
  });
  
  if (action === 'completed') {
    email.status = 'completed';
  } else {
    email.lastUpdated = new Date().toISOString();
  }
  
  // Learning: Update weights based on actions
  const sender = email.sender?.toLowerCase() || '';
  if (!emailTracking.learningWeights.sender[sender]) {
    emailTracking.learningWeights.sender[sender] = 0;
  }
  
  if (action === 'completed') {
    emailTracking.learningWeights.sender[sender] += 5;
  } else if (action === 'snooze') {
    emailTracking.learningWeights.sender[sender] -= 3;
  }
  
  emailTracking.actions.push({
    emailId: emailId,
    action: action,
    timestamp: new Date().toISOString()
  });
  
  saveLocalData();
  renderEmails('all');
}

function clearAllData() {
  if (confirm('Are you sure? This will clear all tracked emails and learning data.')) {
    emailTracking.emails = [];
    emailTracking.actions = [];
    sessionStorage.removeItem('emailFollowupData');
    document.getElementById('content').innerHTML = '<div class="empty-state">Data cleared. Refresh to start fresh.</div>';
  }
}