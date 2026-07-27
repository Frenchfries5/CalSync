import { NextResponse } from "next/server";
import { requireOperator } from "@/auth";
import { applyOne } from "@/lib/calendar";
import type { AppliedRow, MeetingRow } from "@/lib/types";

export const maxDuration = 300;
export const runtime = "nodejs";

type Incoming = Pick<MeetingRow, "key" | "sourceEventId" | "sourceCalendarId" | "subject">;

export async function POST(request: Request) {
  try {
    await requireOperator();

    const body = (await request.json()) as {
      newHireEmail?: string;
      sourceMailbox?: string;
      attendeeType?: string;
      items?: Incoming[];
    };

    const newHireEmail = (body.newHireEmail || "").trim();
    const sourceMailbox = (body.sourceMailbox || "").trim();
    const attendeeType = body.attendeeType === "optional" ? "optional" : "required";
    const items = (body.items || []).filter(
      (item) => item && typeof item.sourceEventId === "string" && item.sourceEventId,
    );

    if (!newHireEmail || !sourceMailbox) {
      return NextResponse.json(
        { error: "Missing new hire email or source mailbox." },
        { status: 400 },
      );
    }
    if (items.length === 0) {
      return NextResponse.json({ error: "No meetings selected." }, { status: 400 });
    }

    // Sequential on purpose: each of these sends real meeting mail, and pushing
    // Exchange hard on one mailbox invites throttling (429).
    const results: AppliedRow[] = [];
    for (const item of items) {
      try {
        results.push(await applyOne(item, sourceMailbox, newHireEmail, attendeeType));
      } catch (error) {
        results.push({
          key: item.key,
          sourceEventId: item.sourceEventId,
          sourceCalendarId: item.sourceCalendarId,
          subject: item.subject,
          organizerEmail: "",
          organizerName: "",
          isRecurring: false,
          method: "direct",
          isAllDay: false,
          status: "error",
          detail: error instanceof Error ? error.message : String(error),
        });
      }
    }

    return NextResponse.json({ results });
  } catch (error) {
    if (error instanceof Response) return error;
    return NextResponse.json(
      { error: error instanceof Error ? error.message : String(error) },
      { status: 400 },
    );
  }
}
