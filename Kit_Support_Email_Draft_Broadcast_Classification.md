Subject: API question: classify broadcasts as Value vs Sales before send (and retrieve in v4)

Hi Kit team,

I’m building an internal marketing dashboard that pulls real-time data from the Kit v4 API.

We need to classify each broadcast as either:
- Value
- Sales

Today we can see standard broadcast fields (subject, description/internal note, send dates, stats, etc.), but we don’t see a native Value/Sales classification field.

Could you please confirm:

1) Is there any native broadcast-level field in Kit (UI or API) to label/type a broadcast as Value vs Sales before sending?

2) If yes:
- Which endpoint/field should we use?
- Can it be set pre-send and updated post-send?
- Can we filter broadcasts by this field via API?

3) If not:
- Is using `description` (internal note) with a structured prefix (example: `[TYPE:VALUE]`, `[TYPE:SALES]`) the recommended workaround?
- Is `description` reliably available in the broadcast list/stats endpoints for downstream analytics?

4) Do you have any roadmap for first-class broadcast categorization/taxonomy in API v4?

For context, our dashboard tracks Value OR, Sales OR, and Sales CTOR in near real time, so stable broadcast classification is essential.

Thanks a lot for the help.

Best,  
[Your Name]
