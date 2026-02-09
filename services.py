# services.py
from datetime import datetime
from sqlalchemy.orm import Session
from sqlalchemy import select, func
from db import Part

def create_part_with_ot(db: Session, client: str) -> Part:
    year_full = datetime.now().year
    yy = year_full % 100

    max_seq = db.execute(
        select(func.max(Part.seq)).where(Part.year == year_full)
    ).scalar()

    next_seq = int(max_seq or 0) + 1
    ot = f"{yy:02d}OT{next_seq:05d}"

    part = Part(
        ot=ot,
        year=year_full,
        seq=next_seq,
        client=(client or "").strip(),
        is_closed=False,
    )
    db.add(part)
    db.commit()
    db.refresh(part)
    return part
