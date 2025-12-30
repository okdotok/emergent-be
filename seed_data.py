#!/usr/bin/env python3
"""
Seed test data for The Global Urenregistratie
"""
import asyncio
import os
import sys
from datetime import datetime, timezone, timedelta
from motor.motor_asyncio import AsyncIOMotorClient
from dotenv import load_dotenv
from pathlib import Path
import uuid
from passlib.context import CryptContext
from math import radians, sin, cos, sqrt, atan2

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

mongo_url = os.environ['MONGO_URL']

# Ensure MongoDB Atlas connection string has proper parameters
if 'mongodb+srv://' in mongo_url:
    if '?' in mongo_url:
        query_part = mongo_url.split('?')[1]
        if 'retryWrites=' not in query_part:
            separator = '&' if query_part else ''
            mongo_url += f'{separator}retryWrites=true'
        if 'w=' not in query_part:
            mongo_url += '&w=majority'
    else:
        mongo_url += '?retryWrites=true&w=majority'

client = AsyncIOMotorClient(
    mongo_url,
    serverSelectionTimeoutMS=30000,
    connectTimeoutMS=20000,
)
db = client[os.environ['DB_NAME']]

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

def calculate_distance(lat1, lon1, lat2, lon2):
    """Calculate distance in meters using Haversine formula"""
    R = 6371000  # Earth radius in meters
    
    lat1_rad = radians(lat1)
    lat2_rad = radians(lat2)
    delta_lat = radians(lat2 - lat1)
    delta_lon = radians(lon2 - lon1)
    
    a = sin(delta_lat / 2) ** 2 + cos(lat1_rad) * cos(lat2_rad) * sin(delta_lon / 2) ** 2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))
    
    distance = R * c
    return distance

async def seed_data():
    print("🌱 Starting data seeding...")
    
    # 1. Create test invitations
    print("\n📧 Creating test invitations...")
    test_invitations = [
        {"email": "jan.jansen@test.nl"},
        {"email": "dummy+1@test.nl"},
        {"email": "test.user@example.com"},
    ]
    
    for inv_data in test_invitations:
        existing = await db.invitations.find_one({"email": inv_data["email"]})
        if not existing:
            invitation = {
                "id": str(uuid.uuid4()),
                "email": inv_data["email"],
                "token": str(uuid.uuid4()),
                "used": False,
                "created_by": "admin",
                "created_at": datetime.now(timezone.utc).isoformat()
            }
            await db.invitations.insert_one(invitation)
            print(f"  ✅ Created invitation for {inv_data['email']}")
        else:
            print(f"  ⏭️  Invitation for {inv_data['email']} already exists")
    
    # 2. Create test projects with coordinates
    print("\n🏢 Creating test projects with GPS coordinates...")
    # Utrecht centrum coordinates
    utrecht_lat, utrecht_lon = 52.0907, 5.1214
    
    test_projects = [
        {
            "name": "Kantoor Renovatie",
            "company": "Te bepalen",
            "location": "Utrecht Centrum",
            "latitude": utrecht_lat,
            "longitude": utrecht_lon,
            "location_radius": 100.0,
            "description": "Hoofdkantoor renovatie project"
        },
        {
            "name": "Bouwproject Noord",
            "company": "Te bepalen", 
            "location": "Amsterdam Noord",
            "latitude": 52.3702,
            "longitude": 4.9041,
            "location_radius": 150.0,
            "description": "Nieuwbouw project in Amsterdam"
        }
    ]
    
    project_ids = []
    for proj_data in test_projects:
        existing = await db.projects.find_one({"name": proj_data["name"]})
        if not existing:
            project = {
                "id": str(uuid.uuid4()),
                **proj_data,
                "active": True,
                "created_at": datetime.now(timezone.utc).isoformat()
            }
            await db.projects.insert_one(project)
            project_ids.append((project["id"], project["name"], project["latitude"], project["longitude"]))
            print(f"  ✅ Created project: {proj_data['name']}")
        else:
            project_ids.append((existing["id"], existing["name"], existing.get("latitude"), existing.get("longitude")))
            print(f"  ⏭️  Project {proj_data['name']} already exists")
    
    # 3. Create test employee users if they don't exist
    print("\n👥 Creating test employee users...")
    test_employees = [
        {
            "email": "employee1@test.nl",
            "first_name": "Jan",
            "last_name": "Jansen",
            "role": "employee"
        },
        {
            "email": "employee2@test.nl",
            "first_name": "Piet",
            "last_name": "de Vries",
            "role": "employee"
        }
    ]
    
    for emp_data in test_employees:
        existing = await db.users.find_one({"email": emp_data["email"]})
        if not existing:
            employee = {
                "id": str(uuid.uuid4()),
                **emp_data,
                "password": pwd_context.hash("test123"),
                "created_at": datetime.now(timezone.utc).isoformat()
            }
            await db.users.insert_one(employee)
            print(f"  ✅ Created employee: {emp_data['first_name']} {emp_data['last_name']} ({emp_data['email']})")
        else:
            print(f"  ⏭️  Employee {emp_data['email']} already exists")
    
    # Get all users (including admin and employees)
    print("\n👥 Finding all users...")
    users = await db.users.find({}, {"_id": 0}).to_list(100)
    print(f"  ✅ Found {len(users)} users")
    
    # 4. Create test time entries
    print("\n⏰ Creating test time entries...")
    
    # Get project for entries
    if not project_ids:
        print("  ⚠️  No projects available. Skipping time entries.")
        return
    
    # Create entries for the past week
    entries_created = 0
    for user in users:
        if user['role'] != 'employee':
            continue  # Skip admin users
        
        for days_ago in range(7, 0, -1):  # Last 7 days
            entry_date = datetime.now(timezone.utc) - timedelta(days=days_ago)
            
            for project_id, project_name, proj_lat, proj_lon in project_ids[:1]:  # Use first project
                # Create 2 entries: one within range, one outside
                for idx, distance_offset in enumerate([0.001, 0.003]):  # ~111m and ~333m
                    clock_in_time = entry_date.replace(hour=8+idx*4, minute=0, second=0, microsecond=0)
                    clock_out_time = clock_in_time + timedelta(hours=4)
                    
                    # Calculate location (within or outside 250m)
                    entry_lat = proj_lat + distance_offset
                    entry_lon = proj_lon + distance_offset
                    
                    distance = calculate_distance(entry_lat, entry_lon, proj_lat, proj_lon)
                    project_match = distance <= 250  # DEFAULT_PROJECT_MATCH_RADIUS
                    
                    location_warning = None
                    if distance > 100:  # project radius
                        location_warning = f"WAARSCHUWING: Locatie afwijking {int(distance)}m (toegestaan: 100m)"
                    
                    entry = {
                        "id": str(uuid.uuid4()),
                        "user_id": user["id"],
                        "user_name": f"{user['first_name']} {user['last_name']}",
                        "project_id": project_id,
                        "project_name": project_name,
                        "company": "Te bepalen",
                        "project_location": "Utrecht Centrum" if project_id == project_ids[0][0] else "Amsterdam Noord",
                        "clock_in_time": clock_in_time.isoformat(),
                        "clock_in_location": {
                            "latitude": entry_lat,
                            "longitude": entry_lon,
                            "accuracy": 10.0
                        },
                        "clock_out_time": clock_out_time.isoformat(),
                        "clock_out_location": {
                            "latitude": entry_lat + 0.0001,
                            "longitude": entry_lon + 0.0001,
                            "accuracy": 10.0
                        },
                        "total_hours": 4.0,
                        "status": "clocked_out",
                        "location_warning": location_warning,
                        "distance_to_project_m": distance,
                        "project_match": project_match,
                        "note": f"Test entry {idx+1}",
                        "created_at": clock_in_time.isoformat()
                    }
                    
                    await db.clock_entries.insert_one(entry)
                    entries_created += 1
    
    print(f"  ✅ Created {entries_created} test time entries")
    print(f"     - Some entries within 250m (project_match=true)")
    print(f"     - Some entries outside 250m (project_match=false)")
    
    print("\n✅ Data seeding complete!")
    print("\nTest data created:")
    print(f"  - {len(test_invitations)} test invitations")
    print(f"  - {len(project_ids)} projects with GPS coordinates")
    print(f"  - {entries_created} time entries (some outside 250m radius)")

if __name__ == "__main__":
    asyncio.run(seed_data())
