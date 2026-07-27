import { NextResponse } from "next/server";
import { requireOperator } from "@/auth";
import { listWeekOccurrences } from "@/lib/calendar";

export const maxDuration = 60;
export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    await requireOperator();

    const body = (await request.json()) as {
      sourceMailbox?: string;
      calendar?: string;
      weekStart?: string;
      weekEnd?: string;
    };

    const sourceMailbox = (body.sourceMailbox || "").trim();
    if (!sourceMailbox || !body.weekStart || !body.weekEnd) {
      return NextResponse.json(
        { error: "Missing sourceMailbox, weekStart or weekEnd." },
        { status: 400 },
      );
    }

    const occurrences = await listWeekOccurrences(
      sourceMailbox,
      (body.calendar || "").trim() || undefined,
      body.weekStart,
      body.weekEnd,
    );
    return NextResponse.json({ occurrences });
  } catch (error) {
    if (error instanceof Response) return error;
    return NextResponse.json(
      { error: error instanceof Error ? error.message : String(error) },
      { status: 400 },
    );
  }
}
