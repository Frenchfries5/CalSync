import { NextResponse } from "next/server";
import { requireOperator } from "@/auth";
import { DEFAULT_MIN_GROUP_MEMBERS, DEFAULT_WINDOW_DAYS, scanMailbox } from "@/lib/calendar";

// Scanning several calendars and resolving each organizer's copy is a lot of
// sequential Graph round-trips. Val Town capped this at ~60s wall clock, which
// is why the original had to batch; here we just ask for the headroom.
export const maxDuration = 120;
export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    await requireOperator();

    const body = (await request.json()) as {
      newHireEmail?: string;
      sourceMailbox?: string;
      referenceEmails?: string[];
      calendar?: string;
      windowDays?: number;
      minGroupMembers?: number;
    };

    const newHireEmail = (body.newHireEmail || "").trim();
    const sourceMailbox = (body.sourceMailbox || "").trim();
    if (!newHireEmail) {
      return NextResponse.json({ error: "Enter the new hire's email." }, { status: 400 });
    }
    if (!sourceMailbox) {
      return NextResponse.json(
        { error: "Enter the mailbox whose calendar you want to mirror." },
        { status: 400 },
      );
    }

    const result = await scanMailbox({
      sourceMailbox,
      newHireEmail,
      referenceEmails: Array.isArray(body.referenceEmails) ? body.referenceEmails : [],
      calendarNameFilter: (body.calendar || "").trim() || undefined,
      windowDays: Number(body.windowDays) || DEFAULT_WINDOW_DAYS,
      minGroupMembers: Number(body.minGroupMembers) || DEFAULT_MIN_GROUP_MEMBERS,
    });

    return NextResponse.json(result);
  } catch (error) {
    if (error instanceof Response) return error;
    return NextResponse.json(
      { error: error instanceof Error ? error.message : String(error) },
      { status: 400 },
    );
  }
}
