/***************
 * CONFIG
 ***************/
var __MEMO = {};
const CFG = {
    SHEETS: {
        USERS: "Users",
        PAGES: "Pages",
        PAGE_REVIEWERS: "PageReviewers",
        TOPICS: "Topics",
        TOPIC_REVIEWS: "TopicReviews",
        ARTICLES: "Articles",
        ARTICLE_REVIEWS: "ArticleReviews",
        AUDIT: "AuditLog",
        SLACK_TEMPLATES: "SlackTemplates",
    },
    TOPIC_STATUSES: {
        UNDER_REVIEW: "Under Review",
        CHANGES: "Changes Requested",
        APPROVED: "Approved",
        DISCARDED: "Discarded",
        REJECTED: "Rejected",
        DRAFT: "Draft",
    },
    ID_YEAR: "2026",
    TOPIC_SEQ_KEY: "TOPIC_SEQ_2026",
    ARTICLE_STATUSES: {
        UNDER_REVIEW: "Under Review",
        CHANGES: "Changes Requested",
        READY: "Ready to Post",
        POSTED: "Posted",
        REJECTED: "Rejected",
    },
    NOTIFICATION_TYPES: {
        TOPIC_REVIEW: "topic_slack_channel",
        ARTICLE_REVIEW: "article_slack_channel",
        PUBLISHER: "publisher_slack_channel",
    }
};

/***************
 * ONE-TIME SETUP (IMPORTANT)
 * Run this once from Apps Script editor: setup_storeSpreadsheetId
 ***************/
function setup_storeSpreadsheetId() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("No active spreadsheet found. Open the spreadsheet-bound Apps Script project.");
    PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ss.getId());
}

/***************
 * WEB APP ENTRY
 ***************/
function doGet() {
    return HtmlService.createHtmlOutputFromFile("index")
        .setTitle("Social Content Approval Portal")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***************
 * API: AUTH/CONTEXT
 ***************/
function api_getContext(clientEmail) {
    try {
        // Consumer Gmail often can't be auto-detected in web apps.
        // We prioritize the client-provided email to allow "switching users" (masquerading) for the MVP,
        // especially since session detection is flaky for personal accounts.
        const sessionEmail = (Session.getActiveUser && Session.getActiveUser().getEmail()) || "";
        const email = ((clientEmail && clientEmail.trim()) || sessionEmail || "").trim().toLowerCase();

        if (!email) {
            return { ok: false, code: "NO_EMAIL", message: "No email detected. Please enter your email." };
        }

        const user = getUserByEmail_(email);
        if (!user || user.is_active !== "Y") {
            return { ok: false, code: "NOT_ALLOWED", message: "Your email is not allowlisted (Users tab) or not active." };
        }

        return {
            ok: true,
            email,
            name: user.name || email,
            roles: {
                is_super_admin: user.is_super_admin === "Y",
                is_author: user.is_author === "Y",
                is_article_reviewer: user.is_article_reviewer === "Y",
                is_publisher: user.is_publisher === "Y",
            },
            ooo: {
                is_out_of_office: user.is_out_of_office === "Y",
                ooo_until: user.ooo_until || ""
            },
            pages: listPages_(),
        };
    } catch (error) {
        // Catch any server-side errors and return them to the client
        return {
            ok: false,
            code: "SERVER_ERROR",
            message: "Server error: " + error.message,
            stack: error.stack
        };
    }
}

/***************
 * API: TOPICS
 ***************/
function api_submitTopic(payload) {
    const email = mustAllow_(payload.email);
    const title = (payload.topic_title || "").trim();
    const page_id = (payload.page_id || "").trim();
    const notes = (payload.notes || "").trim();

    if (!title) throw new Error("Topic title is required.");
    if (!page_id) throw new Error("Page is required.");

    const pages = listPages_();
    if (!pages.find((p) => p.page_id === page_id)) throw new Error("Unknown page_id.");

    // Use LockService to prevent duplicate submissions
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    const topic_id = nextId_(page_id, "SEQ_" + page_id + "_" + CFG.ID_YEAR);
    const now = new Date();

    const approvals_required = getApprovalsRequired_(page_id);

    const topicsSheet = sheet_(CFG.SHEETS.TOPICS);
    const cols = headers_(topicsSheet);

    const rowObj = {
        topic_id,
        topic_title: title,
        page_id,
        author_email: email,
        status: CFG.TOPIC_STATUSES.UNDER_REVIEW,
        cycle_no: 1,
        submitted_at: now,
        last_status_changed_at: now,
        approved_at: "",
        discarded_at: "",
        approvals_required,
        approvals_count: 0,
        has_objection: "N",
        approvals_remaining: approvals_required,
        awaiting_minutes: "",
        notes,
        content_doc_url: (payload.content_doc_url || "").trim(), // New field
    };

    appendRowByHeaders_(topicsSheet, cols, rowObj);
    lock.releaseLock();

    audit_("TOPIC", topic_id, "SUBMIT_TOPIC", "", rowObj.status, email, now, "");

    const slackMsg = getSlackMessage_("TOPIC_SUBMIT", {
        topic_id,
        title,
        author: email
    });
    sendSlackNotification_(page_id, CFG.NOTIFICATION_TYPES.TOPIC_REVIEW, slackMsg);

    return { ok: true, topic_id };
}

function api_listMyTopics(payload) {
    const email = mustAllow_(payload.email);
    const topics = listTopics_({ author_email: email }).map(t => {
        t.author_name = getUserByEmail_(t.author_email)?.name || t.author_email;
        return t;
    });
    const articles = listArticles_({ author_email: email }).map(a => {
        a.author_name = getUserByEmail_(a.author_email)?.name || a.author_email;
        a.reviewer_name = a.assigned_reviewer_email ? (getUserByEmail_(a.assigned_reviewer_email)?.name || a.assigned_reviewer_email) : "";
        return a;
    });
    return { ok: true, topics, articles };
}

function api_listTopicQueue(payload) {
    const email = mustAllow_(payload.email);
    let topics = listTopics_({ status: CFG.TOPIC_STATUSES.UNDER_REVIEW }).filter(
        (t) => (t.author_email || "").toLowerCase() !== email.toLowerCase()
    ).map(t => {
        t.author_name = getUserByEmail_(t.author_email)?.name || t.author_email;
        return t;
    });

    // Filter out topics already reviewed by this user in the current cycle
    const reviewsSh = sheet_(CFG.SHEETS.TOPIC_REVIEWS);
    const reviewsCols = headers_(reviewsSh);
    const lastRow = reviewsSh.getLastRow();

    if (lastRow >= 2) {
        const reviewData = reviewsSh.getRange(2, 1, lastRow - 1, reviewsSh.getLastColumn()).getValues();
        const myReviews = new Set();

        reviewData.forEach(r => {
            const rEmail = String(reviewsCols["reviewer_email"] ? r[reviewsCols["reviewer_email"] - 1] : "").trim().toLowerCase();
            const rTopicId = String(reviewsCols["topic_id"] ? r[reviewsCols["topic_id"] - 1] : "").trim();
            const rCycle = Number(reviewsCols["cycle_no"] ? r[reviewsCols["cycle_no"] - 1] : 0);

            if (rEmail === email) {
                myReviews.add(`${rTopicId}_${rCycle}`);
            }
        });

        topics = topics.filter(t => !myReviews.has(`${t.topic_id}_${t.cycle_no}`));
    }

    return { ok: true, topics };
}

function api_topicApprove(payload) {
    const email = mustAllow_(payload.email);
    const topic_id = (payload.topic_id || "").trim();
    const comment = (payload.comment || "").trim();

    const t = getTopic_(topic_id);
    if (!t) throw new Error("Topic not found.");
    ensureTopicReviewEligible_(email, t);

    writeTopicReview_(topic_id, Number(t.cycle_no), email, "APPROVE", comment);
    audit_("TOPIC", topic_id, "TOPIC_APPROVE", t.status, t.status, email, new Date(), comment);

    evaluateTopic_(topic_id);
    return { ok: true };
}

function api_topicRequestEdit(payload) {
    const email = mustAllow_(payload.email);
    const topic_id = (payload.topic_id || "").trim();
    const comment = (payload.comment || "").trim();
    if (!comment) throw new Error("Comment is required when requesting edits.");

    const t = getTopic_(topic_id);
    if (!t) throw new Error("Topic not found.");
    ensureTopicReviewEligible_(email, t);

    writeTopicReview_(topic_id, Number(t.cycle_no), email, "EDIT", comment);
    setTopicStatus_(topic_id, CFG.TOPIC_STATUSES.CHANGES, email, "REQUEST EDIT: " + comment);

    const slackMsg = getSlackMessage_("TOPIC_OBJECT", {
        topic_id,
        author: t.author_email,
        reviewer: email,
        comment
    });
    sendSlackNotification_(t.page_id, CFG.NOTIFICATION_TYPES.TOPIC_REVIEW, slackMsg);

    return { ok: true };
}

function api_topicReject(payload) {
    const email = mustAllow_(payload.email);
    const topic_id = (payload.topic_id || "").trim();
    const comment = (payload.comment || "").trim();
    if (!comment) throw new Error("Comment is required when rejecting.");

    const t = getTopic_(topic_id);
    if (!t) throw new Error("Topic not found.");
    ensureTopicReviewEligible_(email, t);

    writeTopicReview_(topic_id, Number(t.cycle_no), email, "REJECT", comment);
    setTopicStatus_(topic_id, CFG.TOPIC_STATUSES.REJECTED, email, "REJECT: " + comment);

    const slackMsg = getSlackMessage_("TOPIC_REJECT", {
        topic_id,
        author: t.author_email,
        reviewer: email,
        comment
    });
    sendSlackNotification_(t.page_id, CFG.NOTIFICATION_TYPES.TOPIC_REVIEW, slackMsg);

    return { ok: true };
}

function api_getTopicHistory(payload) {
    mustAllow_(payload.email);
    const topic_id = (payload.topic_id || "").trim();
    const t = getTopic_(topic_id);
    if (!t) throw new Error("Topic not found.");
    const reviews = listTopicReviews_(topic_id);
    return { ok: true, topic: t, reviews };
}

function api_topicResubmit(payload) {
    const email = mustAllow_(payload.email);
    const topic_id = (payload.topic_id || "").trim();
    const newTitle = (payload.topic_title || "").trim();
    const newNotes = (payload.notes || "").trim();

    const t = getTopic_(topic_id);
    if (!t) throw new Error("Topic not found.");
    if (t.author_email.toLowerCase() !== email) throw new Error("Only the author can resubmit this topic.");

    if (t.status !== CFG.TOPIC_STATUSES.CHANGES) {
        throw new Error("Topic must be in 'Changes Requested' status to resubmit.");
    }

    const nextCycle = Number(t.cycle_no) + 1;
    const now = new Date();
    const approvals_required = getApprovalsRequired_(t.page_id);

    const updates = {
        cycle_no: nextCycle,
        status: CFG.TOPIC_STATUSES.UNDER_REVIEW,
        submitted_at: now, // Reset time for wait calculations
        last_status_changed_at: now,
        approvals_count: 0,
        has_objection: "N",
        approvals_remaining: approvals_required,
        approved_at: "", // Clear specific outcomes
        discarded_at: ""
    };

    if (newTitle) updates.topic_title = newTitle;
    if (newNotes) updates.notes = newNotes;

    updateTopicFields_(topic_id, updates);
    audit_("TOPIC", topic_id, "RESUBMIT", t.status, CFG.TOPIC_STATUSES.UNDER_REVIEW, email, now, `Cycle ${nextCycle}. Notes: ${newNotes}`);

    return { ok: true };
}

function api_topicDiscard(payload) {
    const email = mustAllow_(payload.email);
    const topic_id = (payload.topic_id || "").trim();

    const t = getTopic_(topic_id);
    if (!t) throw new Error("Topic not found.");
    if (t.author_email.toLowerCase() !== email) throw new Error("Only the author can discard this topic.");

    if (t.status === CFG.TOPIC_STATUSES.APPROVED) {
        throw new Error("Cannot discard an approved topic.");
    }
    if (t.status === CFG.TOPIC_STATUSES.DISCARDED) {
        throw new Error("Topic is already discarded.");
    }

    const now = new Date();
    updateTopicFields_(topic_id, {
        status: CFG.TOPIC_STATUSES.DISCARDED,
        last_status_changed_at: now,
        discarded_at: now
    });
    audit_("TOPIC", topic_id, "DISCARD", t.status, CFG.TOPIC_STATUSES.DISCARDED, email, now, "User requested discard");

    return { ok: true };
}

/***************
 * API: ARTICLES
 ***************/
function api_submitArticle(payload) {
    const email = mustAllow_(payload.email);
    const topic_id = (payload.topic_id || "").trim();
    const content_url = (payload.content_doc_url || "").trim();
    const notes = (payload.notes_to_reviewer || "").trim();

    if (!content_url) throw new Error("Content URL is required.");

    const t = getTopic_(topic_id);
    if (!t) throw new Error("Topic not found.");
    if (t.status !== CFG.TOPIC_STATUSES.APPROVED) throw new Error("Topic must be approved to submit article.");
    if (t.author_email.toLowerCase() !== email) throw new Error("Only topic author can submit article.");

    // Prevent duplicate article submissions
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    // Reuse topic ID for article ID as requested
    const article_id = topic_id;
    const now = new Date();

    const sh = sheet_(CFG.SHEETS.ARTICLES);
    const cols = headers_(sh);

    const rowObj = {
        article_id,
        topic_id,
        page_id: t.page_id,
        author_email: email,
        status: CFG.ARTICLE_STATUSES.UNDER_REVIEW,
        submitted_at: now,
        last_status_changed_at: now,
        assigned_reviewer_email: "",
        claimed_at: "",
        content_doc_url: content_url,
        notes_to_reviewer: notes,
        posted_at: "",
        posted_url: ""
    };

    appendRowByHeaders_(sh, cols, rowObj);
    lock.releaseLock();

    audit_("ARTICLE", article_id, "SUBMIT_ARTICLE", "", rowObj.status, email, now, "");

    // Find all active reviewers for this page to mention them
    const mentors = getMentionsForPage_(t.page_id);
    const slackMsg = getSlackMessage_("ARTICLE_SUBMIT", {
        article_id,
        topic_id,
        author: email,
        reviewers: mentors
    });
    sendSlackNotification_(t.page_id, CFG.NOTIFICATION_TYPES.ARTICLE_REVIEW, slackMsg);

    return { ok: true, article_id };
}

function api_listArticles(payload) {
    const email = mustAllow_(payload.email);
    const user = getUserByEmail_(email) || {};
    const isSuperAdmin = user.is_super_admin === "Y";

    const articles = listArticles_({});
    const topics = listTopics_({});
    const topicMap = {};
    topics.forEach(t => {
        topicMap[t.topic_id] = t.topic_title;
    });

    const queue = [];
    const publisherQueue = [];

    articles.forEach(a => {
        // Build resolved names for both queues
        a.author_name = getUserByEmail_(a.author_email)?.name || a.author_email;
        a.reviewer_name = a.assigned_reviewer_email ? (getUserByEmail_(a.assigned_reviewer_email)?.name || a.assigned_reviewer_email) : "";
        a.topic_title = topicMap[a.topic_id] || ("Topic " + a.topic_id);

        // Article Review Queue Logic
        if (a.status === CFG.ARTICLE_STATUSES.UNDER_REVIEW || a.status === CFG.ARTICLE_STATUSES.CHANGES) {
            if (isSuperAdmin) {
                queue.push(a);
                return;
            }

            // EXCLUDE OWN ARTICLES from review
            if (a.author_email.toLowerCase() === email.toLowerCase()) {
                return;
            }

            // If unassigned, anyone eligible can see
            if (!a.assigned_reviewer_email) {
                queue.push(a);
            } else if (a.assigned_reviewer_email.toLowerCase() === email.toLowerCase()) {
                // Assigned to me
                queue.push(a);
            }
        }

        // Publisher Queue Logic
        if (a.status === CFG.ARTICLE_STATUSES.READY) {
            publisherQueue.push(a);
        }
    });

    return { ok: true, queue, publisherQueue };
}

function api_getHistory(payload) {
    try {
        const email = mustAllow_(payload.email);
        const object_type = (payload.object_type || "").trim().toUpperCase(); // TOPIC or ARTICLE
        const object_id = (payload.object_id || "").trim();

        if (!object_id) throw new Error("Missing ID.");

        const sh = sheet_(CFG.SHEETS.AUDIT);
        const data = sh.getDataRange().getValues();
        const cols = headers_(sh);

        const logs = [];
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const rowType = String(row[cols["object_type"] - 1] || "").trim().toUpperCase();
            const rowId = String(row[(cols["object_id"] || 0) - 1] || "").trim();

            if (rowType === object_type && rowId === object_id) {
                const actorEmail = String(row[(cols["actor_email"] || 0) - 1] || "").trim();
                const rawTs = row[(cols["timestamp"] || 0) - 1];
                let ts = "";
                if (rawTs instanceof Date) {
                    ts = rawTs.toISOString();
                } else if (rawTs) {
                    ts = String(rawTs);
                }

                logs.push({
                    timestamp: ts,
                    action: String(row[(cols["action"] || 0) - 1] || ""),
                    from_status: String(row[(cols["from_status"] || 0) - 1] || ""),
                    to_status: String(row[(cols["to_status"] || 0) - 1] || ""),
                    actor_email: actorEmail,
                    actor_name: getUserByEmail_(actorEmail)?.name || actorEmail,
                    notes: String(row[(cols["notes"] || 0) - 1] || "")
                });
            }
        }

        // Return newest first
        return { ok: true, logs: logs.reverse() };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

function api_claimArticle(payload) {
    const email = mustAllow_(payload.email);
    const article_id = (payload.article_id || "").trim();

    const a = getArticle_(article_id);
    if (!a) throw new Error("Article not found.");
    if (a.status !== CFG.ARTICLE_STATUSES.UNDER_REVIEW) throw new Error("Article not eligible for claim.");
    if (a.assigned_reviewer_email) throw new Error("Article already claimed.");
    if (a.author_email.toLowerCase() === email) throw new Error("Cannot review your own article.");

    // Check Pool
    const pool = sheet_(CFG.SHEETS.PAGE_REVIEWERS).getDataRange().getValues();
    const colsPool = headers_(sheet_(CFG.SHEETS.PAGE_REVIEWERS));
    const isEligible = pool.some(r =>
        String(r[colsPool["page_id"] - 1]).trim() === a.page_id &&
        String(r[colsPool["reviewer_email"] - 1]).trim().toLowerCase() === email &&
        String(r[colsPool["is_active"] - 1]).trim().toUpperCase() === "Y"
    );

    // Check if user is super admin
    const user = getUserByEmail_(email);
    const isAdmin = user && user.is_super_admin === "Y";

    if (!isEligible && !isAdmin) {
        throw new Error("You are not in the reviewer pool for this page.");
    }

    const now = new Date();
    updateArticleFields_(article_id, {
        assigned_reviewer_email: email,
        claimed_at: now
    });
    audit_("ARTICLE", article_id, "CLAIM", a.status, a.status, email, now, "");

    // ARTICLE_CLAIM removed as per user request
    return { ok: true };
}

function api_reviewArticle(payload) {
    const email = mustAllow_(payload.email);
    const article_id = (payload.article_id || "").trim();
    const decision = (payload.decision || "").trim(); // APPROVE, CHANGES, REJECT
    const comment = (payload.comment || "").trim();

    const a = getArticle_(article_id);
    if (!a) throw new Error("Article not found.");
    if (a.assigned_reviewer_email.toLowerCase() !== email) throw new Error("You are not the assigned reviewer.");

    const now = new Date();
    let newStatus = a.status;

    if (decision === "APPROVE") {
        newStatus = CFG.ARTICLE_STATUSES.READY;
    } else if (decision === "CHANGES") {
        if (!comment) throw new Error("Comment required for changes.");
        newStatus = CFG.ARTICLE_STATUSES.CHANGES;
    } else if (decision === "REJECT") {
        if (!comment) throw new Error("Comment required for rejection.");
        newStatus = CFG.ARTICLE_STATUSES.REJECTED;
    } else {
        throw new Error("Invalid decision.");
    }

    // Write review log
    const shReviews = sheet_(CFG.SHEETS.ARTICLE_REVIEWS);
    const colsRev = headers_(shReviews);
    appendRowByHeaders_(shReviews, colsRev, {
        review_id: nextId_("AR", "AR_SEQ_" + CFG.ID_YEAR),
        article_id,
        reviewer_email: email,
        decision,
        comment,
        decided_at: now
    });

    // Update status
    updateArticleFields_(article_id, {
        status: newStatus,
        last_status_changed_at: now,
        // If changes requested, unassign?? Requirement says so? Checking req.txt...
        // Req 5.2.6: "On resubmit: clear assigned_reviewer_email...". So on REVIEW decision 'Changes', we keep assignment?
        // Usually reviewer keeps it until they reject or author resubmits. Staying assigned for now.
    });

    audit_("ARTICLE", article_id, "REVIEW_" + decision, a.status, newStatus, email, now, comment);

    const templateId = "ARTICLE_" + decision; // ARTICLE_APPROVE, ARTICLE_CHANGES, ARTICLE_REJECT
    const slackMsg = getSlackMessage_(templateId, {
        article_id,
        author: a.author_email,
        reviewer: email,
        comment: comment || ""
    });

    if (slackMsg) {
        sendSlackNotification_(a.page_id, CFG.NOTIFICATION_TYPES.ARTICLE_REVIEW, slackMsg);
    }

    return { ok: true };
}

function api_markPosted(payload) {
    const email = mustAllow_(payload.email);
    const article_id = (payload.article_id || "").trim();
    const posted_url = (payload.posted_url || "").trim();

    const a = getArticle_(article_id);
    if (!a) throw new Error("Article not found.");
    if (a.status !== CFG.ARTICLE_STATUSES.READY) throw new Error("Article not ready to post.");

    const now = new Date();
    updateArticleFields_(article_id, {
        status: CFG.ARTICLE_STATUSES.POSTED,
        posted_at: now,
        posted_url
    });
    audit_("ARTICLE", article_id, "MARK_POSTED", a.status, CFG.ARTICLE_STATUSES.POSTED, email, now, posted_url);

    const slackMsg = getSlackMessage_("ARTICLE_POSTED", {
        article_id,
        posted_url
    });
    sendSlackNotification_(a.page_id, CFG.NOTIFICATION_TYPES.PUBLISHER, slackMsg);

    return { ok: true };
}

function api_articleResubmit(payload) {
    const email = mustAllow_(payload.email);
    const article_id = (payload.article_id || "").trim();
    const content_url = (payload.content_doc_url || "").trim();
    const notes = (payload.notes_to_reviewer || "").trim();
    const image_urls = (payload.image_urls || "").trim();
    const posting_dt = (payload.preferred_posting_datetime || "").trim();
    const language = (payload.content_language || "").trim();

    const a = getArticle_(article_id);
    if (!a) throw new Error("Article not found.");
    if (a.author_email.toLowerCase() !== email) throw new Error("Only author can resubmit.");
    if (a.status !== CFG.ARTICLE_STATUSES.CHANGES) throw new Error("Article must be in 'Changes Requested' to resubmit.");

    const now = new Date();
    const updates = {
        status: CFG.ARTICLE_STATUSES.UNDER_REVIEW,
        last_status_changed_at: now,
        submitted_at: now,
        assigned_reviewer_email: "", // Requirement 5.2.6: Clear assignment
        claimed_at: "",
    };

    if (content_url) updates.content_doc_url = content_url;
    if (notes) updates.notes_to_reviewer = notes;
    if (image_urls !== undefined) updates.image_urls = image_urls;
    if (posting_dt !== undefined) updates.preferred_posting_datetime = posting_dt;
    if (language !== undefined) updates.content_language = language;

    updateArticleFields_(article_id, updates);
    audit_("ARTICLE", article_id, "RESUBMIT", CFG.ARTICLE_STATUSES.CHANGES, CFG.ARTICLE_STATUSES.UNDER_REVIEW, email, now, notes);

    const slackMsg = getSlackMessage_("ARTICLE_RESUBMIT", {
        article_id,
        topic_id: a.topic_id,
        author: email
    });
    sendSlackNotification_(a.page_id, CFG.NOTIFICATION_TYPES.ARTICLE_REVIEW, slackMsg);

    return { ok: true };
}

/***************
 * API: ADMIN / OVERRIDES
 ***************/
function api_adminUnassignArticle(payload) {
    const email = mustAllow_(payload.email);
    const user = getUserByEmail_(email);
    if (user.is_super_admin !== "Y") throw new Error("Super-admin only.");

    const article_id = (payload.article_id || "").trim();
    const a = getArticle_(article_id);
    if (!a) throw new Error("Article not found.");

    const now = new Date();
    updateArticleFields_(article_id, {
        assigned_reviewer_email: "",
        claimed_at: ""
    });
    audit_("ARTICLE", article_id, "UNASSIGN", a.status, a.status, email, now, payload.reason || "Admin unassign");

    return { ok: true };
}

function api_adminOverrideArticle(payload) {
    const email = mustAllow_(payload.email);
    const user = getUserByEmail_(email);
    if (user.is_super_admin !== "Y") throw new Error("Super-admin only.");

    const article_id = (payload.article_id || "").trim();
    const newStatus = (payload.status || "").trim();
    const reason = (payload.reason || "").trim();

    const a = getArticle_(article_id);
    if (!a) throw new Error("Article not found.");
    if (!reason) throw new Error("Reason is required for override.");

    const now = new Date();
    updateArticleFields_(article_id, {
        status: newStatus,
        last_status_changed_at: now
    });
    audit_("ARTICLE", article_id, "OVERRIDE", a.status, newStatus, email, now, reason);

    return { ok: true };
}

function api_adminOverrideTopic(payload) {
    const email = mustAllow_(payload.email);
    const user = getUserByEmail_(email);
    if (user.is_super_admin !== "Y") throw new Error("Super-admin only.");

    const topic_id = (payload.topic_id || "").trim();
    const newStatus = (payload.status || "").trim();
    const reason = (payload.reason || "").trim();

    const t = getTopic_(topic_id);
    if (!t) throw new Error("Topic not found.");
    if (!reason) throw new Error("Reason is required for override.");

    const now = new Date();
    updateTopicFields_(topic_id, {
        status: newStatus,
        last_status_changed_at: now
    });
    audit_("TOPIC", topic_id, "OVERRIDE", t.status, newStatus, email, now, reason);

    return { ok: true };
}

/***************
 * DIAGNOSTICS
 ***************/
function api_debug(payload) {
    try {
        // do not require allowlist for debug? still require, so only allowlisted users can see.
        mustAllow_(payload.email);

        const ss = getSpreadsheet_();
        const out = {
            spreadsheet_id: ss.getId(),
            spreadsheet_name: ss.getName(),
            tabs_found: ss.getSheets().map(s => s.getName()),
        };

        const topicsSh = ss.getSheetByName(CFG.SHEETS.TOPICS);
        if (!topicsSh) {
            out.topics_error = `Missing tab: ${CFG.SHEETS.TOPICS}`;
            return { ok: true, debug: out };
        }

        const lastRow = topicsSh.getLastRow();
        const lastCol = topicsSh.getLastColumn();
        out.topics_lastRow = lastRow;
        out.topics_lastCol = lastCol;

        if (lastCol < 1) {
            out.topics_error = "Sheet exists but has 0 columns. Headers are required.";
            out.topics_expected_headers = ["topic_id", "topic_title", "page_id", "author_email", "status", "cycle_no", "submitted_at", "last_status_changed_at", "approved_at", "discarded_at", "approvals_required", "approvals_count", "has_objection", "approvals_remaining", "awaiting_minutes", "notes"];
            return { ok: true, debug: out };
        }

        const rawHeaders = topicsSh.getRange(1, 1, 1, lastCol).getValues()[0];
        out.topics_headers_raw = rawHeaders;
        out.topics_headers_normalized = rawHeaders.map(h => normalizeHeader_(h));

        // Show first 5 data rows (selected columns)
        const cols = headers_(topicsSh);
        out.topics_headerMap = cols;

        const sample = [];
        if (lastRow >= 2) {
            const rows = topicsSh.getRange(2, 1, Math.min(5, lastRow - 1), lastCol).getValues();
            rows.forEach(r => {
                sample.push({
                    topic_id: cols["topic_id"] ? String(r[cols["topic_id"] - 1]) : "(missing col)",
                    author_email: cols["author_email"] ? String(r[cols["author_email"] - 1]) : "(missing col)",
                    status: cols["status"] ? String(r[cols["status"] - 1]) : "(missing col)",
                    submitted_at: cols["submitted_at"] ? String(r[cols["submitted_at"] - 1]) : "(missing col)",
                });
            });
        }
        out.topics_sample = sample;

        // What listTopics_ sees
        out.listTopics_total = listTopics_({}).length;
        out.listTopics_queue_total = listTopics_({ status: CFG.TOPIC_STATUSES.UNDER_REVIEW }).length;

        // Preview what api_listTopicQueue returns for this user
        const queueCandidates = listTopics_({ status: CFG.TOPIC_STATUSES.UNDER_REVIEW });
        const filteredQueue = queueCandidates.filter(t => (t.author_email || "").toLowerCase() !== payload.email.toLowerCase());
        out.user_email_seen = payload.email;
        out.queue_preview_count = filteredQueue.length;
        out.queue_preview_ids = filteredQueue.map(t => t.topic_id);

        return { ok: true, debug: out };
    } catch (e) {
        return { ok: false, message: "Debug failed: " + e.message + "\n" + e.stack };
    }
}

/***************
 * TOPIC EVALUATION
 ***************/
function evaluateTopic_(topic_id) {
    const t = getTopic_(topic_id);
    if (!t) return;
    if (t.status !== CFG.TOPIC_STATUSES.UNDER_REVIEW) return;

    const cycle = Number(t.cycle_no);
    const approvals_required = getApprovalsRequired_(t.page_id);

    const reviews = listTopicReviews_(topic_id).filter((r) => Number(r.cycle_no) === cycle);

    const latestByReviewer = {};
    reviews.forEach((r) => {
        const key = (r.reviewer_email || "").toLowerCase();
        const ts = new Date(r.decided_at).getTime();
        if (!latestByReviewer[key] || ts > latestByReviewer[key].ts) {
            latestByReviewer[key] = { decision: r.decision, ts };
        }
    });

    const decisions = Object.values(latestByReviewer).map((x) => x.decision);
    const hasObjection = decisions.includes("OBJECT");
    const approvals = decisions.filter((d) => d === "APPROVE").length;

    if (hasObjection) {
        setTopicStatus_(topic_id, CFG.TOPIC_STATUSES.CHANGES, "system", "Auto: objection present");
        return;
    }

    if (approvals >= approvals_required) {
        setTopicStatus_(topic_id, CFG.TOPIC_STATUSES.APPROVED, "system", `Auto: approvals ${approvals}/${approvals_required}`);

        const slackMsg = getSlackMessage_("TOPIC_APPROVE", {
            topic_id,
            title: t.topic_title,
            author: t.author_email
        });
        sendSlackNotification_(t.page_id, CFG.NOTIFICATION_TYPES.TOPIC_REVIEW, slackMsg);
        return;
    }

    updateTopicFields_(topic_id, {
        approvals_required,
        approvals_count: approvals,
        has_objection: "N",
        approvals_remaining: Math.max(approvals_required - approvals, 0),
    });
}

function setTopicStatus_(topic_id, newStatus, actor, notes) {
    const t = getTopic_(topic_id);
    if (!t) return;
    const now = new Date();

    const update = {
        status: newStatus,
        last_status_changed_at: now,
    };

    if (newStatus === CFG.TOPIC_STATUSES.APPROVED) update.approved_at = now;
    if (newStatus === CFG.TOPIC_STATUSES.CHANGES) update.has_objection = "Y";

    updateTopicFields_(topic_id, update);
    audit_("TOPIC", topic_id, "SET_STATUS", t.status, newStatus, actor, now, notes || "");
}

/***************
 * ELIGIBILITY
 ***************/
function ensureTopicReviewEligible_(reviewerEmail, topic) {
    if ((topic.author_email || "").toLowerCase() === reviewerEmail) {
        throw new Error("You cannot review your own topic.");
    }
}

/***************
 * STORAGE HELPERS
 ***************/
// __MEMO moved to top

function getSpreadsheet_() {
    if (__MEMO.ss) return __MEMO.ss;
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty("SPREADSHEET_ID");
    if (id) {
        __MEMO.ss = SpreadsheetApp.openById(id);
    } else {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        if (!ss) throw new Error("No active spreadsheet. Run setup_storeSpreadsheetId once.");
        __MEMO.ss = ss;
    }
    return __MEMO.ss;
}

function sheet_(name) {
    const key = "sheet_" + name;
    if (__MEMO[key]) return __MEMO[key];
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(name);
    if (!sh) throw new Error(`Missing sheet/tab: ${name}`);
    __MEMO[key] = sh;
    return sh;
}

function getCache_(key, fetcher, ttlSeconds = 600) {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    if (cached) {
        try {
            return JSON.parse(cached);
        } catch (e) {
            console.error("Cache parse failed for " + key);
        }
    }
    const val = fetcher();
    if (val !== null && val !== undefined) {
        cache.put(key, JSON.stringify(val), ttlSeconds);
    }
    return val;
}

function normalizeHeader_(h) {
    return String(h || "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, "_")
        .replace(/[^a-z0-9_]/g, "");
}

function headers_(sheet) {
    if (!sheet) return {};
    const key = "headers_" + sheet.getName();
    if (__MEMO[key]) return __MEMO[key];

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) throw new Error(`Sheet ${sheet.getName()} has no columns.`);
    const vals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const map = {};
    vals.forEach((raw, i) => {
        const key = normalizeHeader_(raw);
        if (key && !map[key]) map[key] = i + 1;
    });
    __MEMO[key] = map;
    return map;
}

function appendRowByHeaders_(sheet, headerMap, obj) {
    const lastCol = sheet.getLastColumn();
    const row = new Array(lastCol).fill("");
    Object.keys(headerMap).forEach((hKey) => {
        const idx0 = headerMap[hKey] - 1;
        row[idx0] = obj[hKey] === undefined ? "" : obj[hKey];
    });
    sheet.appendRow(row);
}

function findRowByValue_(sheet, headerMap, keyHeader, value) {
    const key = normalizeHeader_(keyHeader);
    const col = headerMap[key];
    if (!col) throw new Error(`Missing column '${keyHeader}' (normalized: '${key}') in ${sheet.getName()}`);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return -1;
    const values = sheet.getRange(2, col, lastRow - 1, 1).getValues().flat();
    const target = String(value).trim();
    for (let i = 0; i < values.length; i++) {
        if (String(values[i]).trim() === target) return i + 2;
    }
    return -1;
}

function updateRowByHeaders_(sheet, headerMap, rowIndex, updates) {
    const lastCol = sheet.getLastColumn();
    const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    Object.keys(updates).forEach((k) => {
        const col = headerMap[normalizeHeader_(k)];
        if (col) row[col - 1] = updates[k];
    });
    sheet.getRange(rowIndex, 1, 1, lastCol).setValues([row]);
}

function nextId_(prefix, propKey) {
    const lock = LockService.getScriptLock();
    // nextId_ level lock is good, but submission APIs now have their own locks too.
    lock.waitLock(10000);
    try {
        const props = PropertiesService.getScriptProperties();
        const current = Number(props.getProperty(propKey) || "0");
        const next = current + 1;
        props.setProperty(propKey, String(next));
        // prefix is now Page ID (e.g. IDS)
        return `${prefix}-${CFG.ID_YEAR}-${String(next).padStart(6, "0")}`;
    } finally {
        lock.releaseLock();
    }
}

function audit_(objectType, objectId, action, fromStatus, toStatus, actorEmail, when, notes) {
    const sh = sheet_(CFG.SHEETS.AUDIT);
    const cols = headers_(sh);
    appendRowByHeaders_(sh, cols, {
        audit_id: nextId_("LOG", "AUDIT_SEQ_" + CFG.ID_YEAR),
        object_type: objectType,
        object_id: objectId,
        action,
        from_status: fromStatus,
        to_status: toStatus,
        actor_email: actorEmail,
        timestamp: when || new Date(),
        notes: notes || "",
    });
}

/**
 * Sends a Slack notification based on page settings.
 */
function sendSlackNotification_(page_id, type, message) {
    try {
        const pages = listPages_();
        const p = pages.find(x => String(x.page_id).trim() === String(page_id).trim());
        if (!p) {
            console.warn(`Slack: Page ID "${page_id}" not found in Pages sheet.`);
            return { ok: false, error: "Page ID not found" };
        }

        const webhookUrl = p[type];
        if (!webhookUrl) {
            console.warn(`Slack: No URL found for page "${page_id}" and type "${type}". Check column headers in Pages sheet.`);
            return { ok: false, error: `No URL for type "${type}"` };
        }
        if (!webhookUrl.startsWith("http")) {
            console.warn(`Slack: URL for "${page_id}" is invalid (must start with http).`);
            return { ok: false, error: "Invalid URL format" };
        }

        const res = UrlFetchApp.fetch(webhookUrl, {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify({ text: message }),
            muteHttpExceptions: true
        });

        const code = res.getResponseCode();
        const body = res.getContentText();

        if (code !== 200) {
            console.error(`Slack Error: Received status ${code} - ${body}`);
            return { ok: false, code, body };
        }
        return { ok: true, body };
    } catch (e) {
        console.error(`Slack Exception: ${e.message}`);
        return { ok: false, error: e.message };
    }
}

/**
 * API: Test Slack Integration
 */
function api_testSlack(payload) {
    const email = mustAllow_(payload.email);
    const page_id = (payload.page_id || "").trim();
    const type = (payload.type || CFG.NOTIFICATION_TYPES.TOPIC_REVIEW).trim();

    if (!page_id) throw new Error("Page ID is required for testing.");

    const status = sendSlackNotification_(page_id, type, `🚀 *Slack Test Successful*\nTriggered by: ${email}\nContext: ${type}`);

    if (!status.ok) {
        throw new Error(status.error || `Slack returned ${status.code}: ${status.body}`);
    }

    return {
        ok: true,
        message: "Slack acknowledged the message.",
        details: `Response body: ${status.body}`
    };
}

/**
 * Gets a Slack message from the SlackTemplates sheet.
 * Template IDs use {placeholder} syntax.
 */
function getSlackMessage_(template_id, data) {
    try {
        const fetchTemplates = () => {
            const tempSh = getSpreadsheet_().getSheetByName(CFG.SHEETS.SLACK_TEMPLATES);
            if (!tempSh) return [];
            return tempSh.getDataRange().getValues();
        };

        const vals = getCache_("SLACK_TEMPLATES_DATA", fetchTemplates, 900); // 15 min cache
        if (!vals || vals.length === 0) return "";

        const sh = sheet_(CFG.SHEETS.SLACK_TEMPLATES);
        const cols = headers_(sh);

        let templateText = "";
        for (let i = 1; i < vals.length; i++) {
            if (String(vals[i][cols["template_id"] - 1]).trim() === template_id) {
                templateText = String(vals[i][cols["message_text"] - 1]);
                break;
            }
        }

        if (!templateText) return "";

        // Fix literal newlines from spreadsheet
        templateText = templateText.replace(/\\n/g, "\n");

        // Helper to get slack mention
        const getMention = (email) => {
            if (!email) return "Unknown";
            const user = getUserByEmail_(email);
            return user && user.slack_user_id ? `<@${user.slack_user_id}>` : email;
        };

        // Standard replace for common fields
        if (data.email) data.email = getMention(data.email);
        if (data.author) data.author = getMention(data.author);
        if (data.reviewer) data.reviewer = getMention(data.reviewer);

        // Replace placeholders
        Object.keys(data).forEach(key => {
            const regex = new RegExp("{" + key + "}", "g");
            templateText = templateText.replace(regex, data[key]);
        });

        return templateText;
    } catch (e) {
        console.error("getSlackMessage_ failed: " + e.message);
        return "";
    }
}

/**
 * Returns a space-separated string of Slack mentions for all active reviewers of a page.
 */
function getMentionsForPage_(page_id) {
    try {
        const sh = sheet_(CFG.SHEETS.PAGE_REVIEWERS);
        const vals = sh.getDataRange().getValues();
        const cols = headers_(sh);

        const mentions = [];
        for (let i = 1; i < vals.length; i++) {
            const pId = String(vals[i][cols["page_id"] - 1]).trim();
            const email = String(vals[i][cols["reviewer_email"] - 1]).trim().toLowerCase();
            const isActive = String(vals[i][cols["is_active"] - 1]).trim().toUpperCase() === "Y";

            if (pId === page_id && isActive) {
                const user = getUserByEmail_(email);
                if (user && user.slack_user_id) {
                    mentions.push(`<@${user.slack_user_id}>`);
                } else {
                    mentions.push(email);
                }
            }
        }
        return mentions.join(" ");
    } catch (e) {
        return "";
    }
}

function mustAllow_(emailRaw) {
    const email = (emailRaw || "").trim().toLowerCase();
    if (!email) throw new Error("Missing email.");
    const user = getUserByEmail_(email);
    if (!user || user.is_active !== "Y") throw new Error("Not allowlisted or inactive.");
    return email;
}

/***************
 * USERS / PAGES
 ***************/
function getUserByEmail_(email) {
    if (!email) return null;
    const cleanEmail = email.trim().toLowerCase();

    // Use __MEMO to cache the entire map once per execution
    if (!__MEMO["_users_map"]) {
        const sh = sheet_(CFG.SHEETS.USERS);
        const data = sh.getDataRange().getValues();
        const cols = headers_(sh);
        const map = {};
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const emailIdx = (cols["email"] || 0) - 1;
            if (emailIdx < 0) continue;

            const uEmail = String(row[emailIdx] || "").trim().toLowerCase();
            if (uEmail) {
                map[uEmail] = {
                    email: uEmail,
                    name: String(row[(cols["name"] || 0) - 1] || "").trim(),
                    is_active: String(row[(cols["is_active"] || 0) - 1] || "").trim().toUpperCase(),
                    is_author: String(row[(cols["is_author"] || 0) - 1] || "").trim().toUpperCase(),
                    is_article_reviewer: String(row[(cols["is_article_reviewer"] || 0) - 1] || "").trim().toUpperCase(),
                    is_publisher: String(row[(cols["is_publisher"] || 0) - 1] || "").trim().toUpperCase(),
                    is_super_admin: String(row[(cols["is_super_admin"] || 0) - 1] || "").trim().toUpperCase(),
                    is_out_of_office: String(row[(cols["is_out_of_office"] || 0) - 1] || "").trim().toUpperCase(),
                    ooo_until: row[(cols["ooo_until"] || 0) - 1] || "",
                    slack_user_id: String(row[(cols["slack_user_id"] || 0) - 1] || "").trim(),
                };
            }
        }
        __MEMO["_users_map"] = map;
    }

    return __MEMO["_users_map"][cleanEmail] || null;
}

function listPages_() {
    return getCache_("PAGES_DATA", () => {
        const sh = sheet_(CFG.SHEETS.PAGES);
        const cols = headers_(sh);
        const lastRow = sh.getLastRow();
        if (lastRow < 2) return [];
        const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
        return rows
            .map((r) => ({
                page_id: String(cols["page_id"] ? r[cols["page_id"] - 1] : "").trim(),
                page_name: String(cols["page_name"] ? r[cols["page_name"] - 1] : "").trim(),
                required_topic_approvals: Number(cols["required_topic_approvals"] ? r[cols["required_topic_approvals"] - 1] : 2) || 2,
                topic_slack_channel: String(cols["topic_slack_channel"] ? r[cols["topic_slack_channel"] - 1] : "").trim(),
                article_slack_channel: String(cols["article_slack_channel"] ? r[cols["article_slack_channel"] - 1] : "").trim(),
                publisher_slack_channel: String(cols["publisher_slack_channel"] ? r[cols["publisher_slack_channel"] - 1] : "").trim(),
            }))
            .filter((p) => p.page_id);
    }, 1800); // 30 min cache
}

function getApprovalsRequired_(page_id) {
    const pages = listPages_();
    const p = pages.find((x) => x.page_id === page_id);
    return p ? Number(p.required_topic_approvals || 2) : 2;
}

/***************
 * TOPICS
 ***************/
function listTopics_(filters) {
    const sh = sheet_(CFG.SHEETS.TOPICS);
    const cols = headers_(sh);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return [];
    const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

    const topics = rows
        .map((r) => ({
            topic_id: String(cols["topic_id"] ? r[cols["topic_id"] - 1] : "").trim(),
            topic_title: String(cols["topic_title"] ? r[cols["topic_title"] - 1] : "").trim(),
            page_id: String(cols["page_id"] ? r[cols["page_id"] - 1] : "").trim(),
            author_email: String(cols["author_email"] ? r[cols["author_email"] - 1] : "").trim().toLowerCase(),
            status: String(cols["status"] ? r[cols["status"] - 1] : "").trim(),
            cycle_no: Number(cols["cycle_no"] ? r[cols["cycle_no"] - 1] : 1) || 1,
            submitted_at: cols["submitted_at"] ? String(r[cols["submitted_at"] - 1]) : "",
        }))
        .filter((t) => t.topic_id);

    if (!filters || Object.keys(filters).length === 0) return topics;

    return topics.filter((t) => {
        if (filters.status && t.status !== filters.status) return false;
        if (filters.author_email && t.author_email !== filters.author_email.toLowerCase()) return false;
        return true;
    });
}

function getTopic_(topic_id) {
    const sh = sheet_(CFG.SHEETS.TOPICS);
    const cols = headers_(sh);
    const rowIndex = findRowByValue_(sh, cols, "topic_id", topic_id);
    if (rowIndex === -1) return null;
    const r = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
    return {
        topic_id: String(cols["topic_id"] ? r[cols["topic_id"] - 1] : "").trim(),
        topic_title: String(cols["topic_title"] ? r[cols["topic_title"] - 1] : "").trim(),
        page_id: String(cols["page_id"] ? r[cols["page_id"] - 1] : "").trim(),
        author_email: String(cols["author_email"] ? r[cols["author_email"] - 1] : "").trim().toLowerCase(),
        status: String(cols["status"] ? r[cols["status"] - 1] : "").trim(),
        cycle_no: Number(cols["cycle_no"] ? r[cols["cycle_no"] - 1] : 1) || 1,
    };
}

function updateTopicFields_(topic_id, updates) {
    const sh = sheet_(CFG.SHEETS.TOPICS);
    const cols = headers_(sh);
    const rowIndex = findRowByValue_(sh, cols, "topic_id", topic_id);
    if (rowIndex === -1) throw new Error("Topic not found.");
    updateRowByHeaders_(sh, cols, rowIndex, updates);
}

/***************
 * ARTICLES HELPERS
 ***************/
function listArticles_(filters) {
    const sh = sheet_(CFG.SHEETS.ARTICLES);
    const cols = headers_(sh);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return [];
    const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

    const articles = rows
        .map((r) => ({
            article_id: String(cols["article_id"] ? r[cols["article_id"] - 1] : "").trim(),
            topic_id: String(cols["topic_id"] ? r[cols["topic_id"] - 1] : "").trim(),
            page_id: String(cols["page_id"] ? r[cols["page_id"] - 1] : "").trim(),
            author_email: String(cols["author_email"] ? r[cols["author_email"] - 1] : "").trim().toLowerCase(),
            status: String(cols["status"] ? r[cols["status"] - 1] : "").trim(),
            assigned_reviewer_email: String(cols["assigned_reviewer_email"] ? r[cols["assigned_reviewer_email"] - 1] : "").trim().toLowerCase(),
            content_doc_url: String(cols["content_doc_url"] ? r[cols["content_doc_url"] - 1] : "").trim(),
            notes_to_reviewer: String(cols["notes_to_reviewer"] ? r[cols["notes_to_reviewer"] - 1] : "").trim(),
            posted_url: String(cols["posted_url"] ? r[cols["posted_url"] - 1] : "").trim(),
            submitted_at: cols["submitted_at"] ? String(r[cols["submitted_at"] - 1]) : "",
            posted_at: cols["posted_at"] ? String(r[cols["posted_at"] - 1]) : "",
        }))
        .filter((a) => a.article_id);

    if (!filters || Object.keys(filters).length === 0) return articles;

    return articles.filter((a) => {
        if (filters.status && a.status !== filters.status) return false;
        if (filters.author_email && a.author_email !== filters.author_email.toLowerCase()) return false;
        return true;
    });
}

function getArticle_(article_id) {
    const sh = sheet_(CFG.SHEETS.ARTICLES);
    const cols = headers_(sh);
    const rowIndex = findRowByValue_(sh, cols, "article_id", article_id);
    if (rowIndex === -1) return null;
    const r = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
    return {
        article_id: String(cols["article_id"] ? r[cols["article_id"] - 1] : "").trim(),
        topic_id: String(cols["topic_id"] ? r[cols["topic_id"] - 1] : "").trim(),
        page_id: String(cols["page_id"] ? r[cols["page_id"] - 1] : "").trim(),
        author_email: String(cols["author_email"] ? r[cols["author_email"] - 1] : "").trim().toLowerCase(),
        status: String(cols["status"] ? r[cols["status"] - 1] : "").trim(),
        assigned_reviewer_email: String(cols["assigned_reviewer_email"] ? r[cols["assigned_reviewer_email"] - 1] : "").trim().toLowerCase(),
        content_doc_url: String(cols["content_doc_url"] ? r[cols["content_doc_url"] - 1] : "").trim(),
        notes_to_reviewer: String(cols["notes_to_reviewer"] ? r[cols["notes_to_reviewer"] - 1] : "").trim(),
    };
}

function updateArticleFields_(article_id, updates) {
    const sh = sheet_(CFG.SHEETS.ARTICLES);
    const cols = headers_(sh);
    const rowIndex = findRowByValue_(sh, cols, "article_id", article_id);
    if (rowIndex === -1) throw new Error("Article not found.");
    updateRowByHeaders_(sh, cols, rowIndex, updates);
}

/***************
 * TOPIC REVIEWS
 ***************/
function writeTopicReview_(topic_id, cycle_no, reviewer_email, decision, comment) {
    const sh = sheet_(CFG.SHEETS.TOPIC_REVIEWS);
    const cols = headers_(sh);
    const now = new Date();
    appendRowByHeaders_(sh, cols, {
        review_id: nextId_("TR", "TOPIC_REVIEW_SEQ_" + CFG.ID_YEAR),
        topic_id,
        cycle_no,
        reviewer_email,
        decision,
        comment: comment || "",
        decided_at: now,
    });
}

function listTopicReviews_(topic_id) {
    const sh = sheet_(CFG.SHEETS.TOPIC_REVIEWS);
    const cols = headers_(sh);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return [];
    const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    return rows
        .map((r) => ({
            topic_id: String(cols["topic_id"] ? r[cols["topic_id"] - 1] : "").trim(),
            cycle_no: cols["cycle_no"] ? r[cols["cycle_no"] - 1] : "",
            reviewer_email: String(cols["reviewer_email"] ? r[cols["reviewer_email"] - 1] : "").trim().toLowerCase(),
            decision: String(cols["decision"] ? r[cols["decision"] - 1] : "").trim(),
            comment: String(cols["comment"] ? r[cols["comment"] - 1] : "").trim(),
            decided_at: cols["decided_at"] ? String(r[cols["decided_at"] - 1]) : "",
        }))
        .filter((x) => x.topic_id === topic_id);
}

// Run this function once from the editor to trigger the permissions popup
function auth_TriggerPermissions() {
    UrlFetchApp.fetch("https://www.google.com");
    console.log("Permissions triggered successfully.");
}
