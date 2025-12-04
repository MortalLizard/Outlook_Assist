/// <reference types="office-js" />
/* global Office, fetch */

// taskpane.ts — robust version that avoids TS errors on displayReplyAllForm

class OutlookAiClient {
  constructor(private baseUrl: string = "http://127.0.0.1:5001") {}

  async draftReply(payload: any): Promise<{ subject: string; body: string }> {
    const res = await fetch(`${this.baseUrl}/assist/reply`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!res.ok) throw new Error(`Server error ${res.status}`);
    return res.json();
  }
}

class OutlookAddinApp {
  private client = new OutlookAiClient();

  init() {
    const btn = document.getElementById("btnDraft");
    btn?.addEventListener("click", () => this.draftFromCurrentItem());
  }

  private async getCurrentItemData(): Promise<any> {
    const mbox = Office.context.mailbox;
    const item = mbox.item as any;

    const me = mbox.userProfile;
    const recipient_display_name = me?.displayName || "Me";

    const incoming_sender_name = item.from?.displayName || "";
    const incoming_sender_email = item.from?.emailAddress || "";

    const incoming_subject: string =
      typeof item.subject === "string" ? item.subject : "";

    const incoming_body = await new Promise<string>((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Html, (res: Office.AsyncResult<string>) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value || "");
        else reject(res.error);
      });
    });

    return {
      recipient_display_name,
      incoming_sender_name,
      incoming_sender_email,
      incoming_subject,
      incoming_body,
      greeting_style: "auto",
      signoff_style: "best_regards",
      language: "en",
      tone: "",
      extra: ""
    };
  }

  public async draftFromCurrentItem() {
    const status = document.getElementById("status");
    const setStatus = (t: string) => { if (status) status.textContent = t; };

    try {
      setStatus("Reading item…");
      const payload = await this.getCurrentItemData();

      setStatus("Calling local assistant…");
      const draft = await this.client.draftReply(payload);

      const item = Office.context.mailbox.item as any;

      // Compose mode: write subject + body directly
      if (item?.subject?.setAsync) {
        await new Promise<void>((resolve, reject) => {
          item.subject.setAsync(draft.subject, (r: Office.AsyncResult<void>) =>
            r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)
          );
        });

        await new Promise<void>((resolve, reject) => {
          item.body.setAsync(
            draft.body,
            { coercionType: Office.CoercionType.Html },
            (r: Office.AsyncResult<void>) =>
              r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)
          );
        });

        setStatus("Draft inserted into compose window.");
      } else {
        // READ mode: open a Reply All window with the drafted body.
        // Some @types/office-js versions make TS picky here; cast mailbox to any
        // and use the string overload to avoid type shape mismatches.
        const mboxAny = Office.context.mailbox as any;
        if (typeof mboxAny.displayReplyAllForm === "function") {
          mboxAny.displayReplyAllForm(draft.body); // string overload (HTML)
        } else {
          // Fallback: open a brand new message (lets us control subject too)
          mboxAny.displayNewMessageForm({
            toRecipients: item.to?.map((r: any) => r.emailAddress) || [],
            ccRecipients: item.cc?.map((r: any) => r.emailAddress) || [],
            subject: `Re: ${payload.incoming_subject || ""}`,
            htmlBody: draft.body
          });
        }
        setStatus("Opened reply window with AI draft.");
      }
    } catch (e: any) {
      setStatus(`Error: ${e?.message || e}`);
    }
  }
}

Office.onReady(() => {
  new OutlookAddinApp().init();
});

export { OutlookAiClient, OutlookAddinApp };
