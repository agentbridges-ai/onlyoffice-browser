/*
 * OnlyOffice editor pages register this service worker from the site root.
 * Keep this bridge at /document_editor_service_worker.js so the registration
 * does not fall through to the app HTML fallback, while preserving this app's
 * root service-worker fetch policy.
 */
importScripts('/sw.js');
