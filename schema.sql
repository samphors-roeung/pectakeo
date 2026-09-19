-- ====================================================================
-- Cloudflare D1 Database Schema for PECTAKEO
-- Compatible with SQLite / Cloudflare D1 Serverless Database
-- ====================================================================

-- 1. System Cache Table (for sub-50ms ultra-fast reads of dashboard/reports)
CREATE TABLE IF NOT EXISTS system_cache (
    cache_key TEXT PRIMARY KEY,
    cache_value TEXT NOT NULL,
    updated_at INTEGER NOT NULL,
    expires_at INTEGER NOT NULL
);

-- Index for quick expiration checks
CREATE INDEX IF NOT EXISTS idx_system_cache_expires ON system_cache(expires_at);

-- 2. Daily Voter Registration Entries (Mirror of Google Sheet entries)
CREATE TABLE IF NOT EXISTS daily_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT NOT NULL,          -- Format: YYYY-MM-DD
    district TEXT NOT NULL,            -- e.g. អង្គរបូរី, បាទី, etc.
    commune TEXT NOT NULL,             -- e.g. គោកធ្លក, ចំបក់, etc.
    values_json TEXT NOT NULL,         -- JSON object containing all 24 fields
    note TEXT,
    created_by TEXT,
    created_at INTEGER NOT NULL,
    updated_at INTEGER NOT NULL
);

-- Compound index for fast district and date queries
CREATE INDEX IF NOT EXISTS idx_daily_entries_date_district ON daily_entries(entry_date, district);
CREATE INDEX IF NOT EXISTS idx_daily_entries_commune ON daily_entries(commune);

-- 3. Users Table (Mirror of Google Sheet 'អ្នកប្រើប្រាស់')
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    role TEXT NOT NULL,
    district TEXT,
    commune TEXT,
    status TEXT NOT NULL DEFAULT 'សកម្ម',
    updated_at INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_users_username ON users(username);
CREATE INDEX IF NOT EXISTS idx_users_role ON users(role);

-- 4. Audit Log Table (Tracks data changes and synchronization events)
CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    action TEXT NOT NULL,
    username TEXT,
    details TEXT,
    status TEXT NOT NULL DEFAULT 'SUCCESS',
    timestamp INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_audit_logs_timestamp ON audit_logs(timestamp);
