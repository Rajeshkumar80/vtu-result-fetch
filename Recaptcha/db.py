"""MongoDB persistence layer for the VTU result fetcher.

Reads MONGODB_URI / MONGODB_DB_NAME from the .env file (python-dotenv).
If MongoDB is not configured or unreachable, every function degrades
gracefully (returns empty results / error messages) — the app keeps working.
"""

import os
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

from pymongo import MongoClient
from pymongo.errors import PyMongoError
from bson import ObjectId

MONGODB_URI = os.getenv("MONGODB_URI", "")
MONGODB_DB_NAME = os.getenv("MONGODB_DB_NAME", "vtu_results")

client = None
db = None
status = "disabled"
status_msg = "MongoDB not configured (set MONGODB_URI in .env)"

# Markers that mean the user hasn't filled in the real connection string yet
_PLACEHOLDER_MARKERS = ("PASTE_ROTATED_PASSWORD_HERE", "<db_password>", "<password>", "<username>", "<user>")


def init_db():
    """Try to connect to MongoDB. Never raises — logs status via return value."""
    global client, db, status, status_msg

    if not MONGODB_URI or any(m in MONGODB_URI for m in _PLACEHOLDER_MARKERS):
        status = "disabled"
        status_msg = "MongoDB not configured — paste your real MONGODB_URI into .env"
        return False

    try:
        client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=5000)
        client.admin.command("ping")
        db = client[MONGODB_DB_NAME]
        status = "connected"
        status_msg = f"Connected to MongoDB: {MONGODB_DB_NAME}"
        return True
    except PyMongoError as e:
        client = None
        db = None
        status = "error"
        msg = str(e)
        if "tls" in msg.lower() or "ssl" in msg.lower():
            status_msg = (
                "MongoDB TLS handshake failed. This usually means your Atlas cluster is "
                "PAUSED or the network blocks port 27017. Open Atlas console, make sure "
                "the cluster shows 'Active', and try again."
            )
        else:
            status_msg = f"MongoDB unreachable: {msg[:200]}"
        return False


def is_connected():
    return status == "connected" and client is not None and db is not None


def _batches_coll():
    return db["fetch_batches"] if db is not None else None


def _students_coll():
    return db["student_results"] if db is not None else None


def _jsonable(doc):
    """Convert a Mongo doc into a JSON-safe dict (ObjectId/date -> str)."""
    out = dict(doc)
    if "_id" in out:
        out["_id"] = str(out["_id"])
    if isinstance(out.get("saved_at"), datetime):
        out["saved_at"] = out["saved_at"].isoformat()
    if "batch_id" in out:
        out["batch_id"] = str(out["batch_id"])
    return out


def insert_batch(batch_doc, student_docs):
    """Insert one fetch_batches doc + N student_results docs (same batch_id).
    Returns (batch_id_str, inserted_student_count) or (None, error_msg)."""
    if not is_connected():
        return None, "MongoDB not connected"
    try:
        batch_id = _batches_coll().insert_one(batch_doc).inserted_id
        for s in student_docs:
            s["batch_id"] = batch_id
        if student_docs:
            _students_coll().insert_many(student_docs)
        return str(batch_id), len(student_docs)
    except PyMongoError as e:
        return None, str(e)


def save_batch(batch_doc, student_docs):
    """Insert a new batch, or MERGE into an existing batch with the same
    (year, scheme, semester, department) — never creates a duplicate record.
    Only USNs not already present in that batch are appended.
    Returns (batch_id_str, added_count, was_merged, total_students)
    or (None, error_msg, False, 0)."""
    if not is_connected():
        return None, "MongoDB not connected", False, 0
    try:
        existing = _batches_coll().find_one({
            "year": batch_doc["year"],
            "scheme": batch_doc["scheme"],
            "semester": batch_doc["semester"],
            "department": batch_doc["department"],
        })

        if existing is None:
            batch_id = _batches_coll().insert_one(batch_doc).inserted_id
            for s in student_docs:
                s["batch_id"] = batch_id
            if student_docs:
                _students_coll().insert_many(student_docs)
            return str(batch_id), len(student_docs), False, len(student_docs)

        # ---- merge path: same batch already exists ----
        oid = existing["_id"]
        existing_usns = set(
            s["usn"] for s in _students_coll().find({"batch_id": oid}, {"usn": 1})
        )
        new_docs = [d for d in student_docs if d["usn"] not in existing_usns]
        for s in new_docs:
            s["batch_id"] = oid
        if new_docs:
            _students_coll().insert_many(new_docs)

        total = _students_coll().count_documents({"batch_id": oid})
        merged_subjects = sorted(
            set(existing.get("subjects") or []) | set(batch_doc.get("subjects") or [])
        )
        _batches_coll().update_one(
            {"_id": oid},
            {"$set": {
                "student_count": total,
                "subjects": merged_subjects,
                "saved_at": datetime.utcnow(),
            }},
        )
        return str(oid), len(new_docs), True, total
    except PyMongoError as e:
        return None, str(e), False, 0


def fetch_batches(limit=100):
    """Return saved batches, newest first."""
    if not is_connected():
        return []
    try:
        docs = list(_batches_coll().find().sort("saved_at", -1).limit(limit))
        return [_jsonable(d) for d in docs]
    except PyMongoError as e:
        print(f"[MongoDB] fetch_batches error: {e}")
        return []


def fetch_batch(batch_id):
    """Return one batch doc (JSON-safe) or None."""
    if not is_connected() or not ObjectId.is_valid(batch_id):
        return None
    try:
        doc = _batches_coll().find_one({"_id": ObjectId(batch_id)})
        return _jsonable(doc) if doc else None
    except PyMongoError as e:
        print(f"[MongoDB] fetch_batch error: {e}")
        return None


def fetch_students(batch_id):
    """Return student docs for one batch (JSON-safe) or []."""
    if not is_connected() or not ObjectId.is_valid(batch_id):
        return []
    try:
        docs = list(_students_coll().find({"batch_id": ObjectId(batch_id)}))
        return [_jsonable(d) for d in docs]
    except PyMongoError as e:
        print(f"[MongoDB] fetch_students error: {e}")
        return []


def fetch_batch_with_students(batch_id):
    """Convenience: (batch_doc, student_docs) or (None, None) if not found."""
    batch = fetch_batch(batch_id)
    if batch is None:
        return None, None
    return batch, fetch_students(batch_id)
