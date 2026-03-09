// Creator comment

It is almost instant, so we can assume it is like repository (MVP), but the best scenario it would be to create additional fallback mechanism - after immediate verification we either return - verified and success or success and required signalR  

# EmployeeId Update Propagation Notes

## What We Know

- Updating `employeeId` through Microsoft Graph uses `PATCH /users/{id}`.
- This is not a `202 Accepted` async workflow exposed by Graph.
- A successful response means Graph accepted and completed the write request.
- In practice, an immediate follow-up read often returns the new `employeeId` almost instantly.

## Important Nuance

- Even though the Graph API call is synchronous, Microsoft Entra is still a distributed system.
- That means a successful update does not guarantee that every read path or downstream Microsoft 365 system will reflect the new value at the exact same moment.
- For Graph itself, the value is often visible immediately or within a short time window.
- For downstream systems, propagation can take longer.

## Practical Interpretation

- `PATCH` success means the change was recorded by Microsoft Graph.
- An immediate `GET` is a reasonable way to verify whether the new `employeeId` is already visible through Graph.
- If the immediate `GET` returns the new value, we can treat the change as verified for the Graph/Entra use case.
- If the immediate `GET` does not return the new value yet, we should treat this as a propagation delay rather than a failed write, unless Graph returned an error.

## Recommended UX

### Option 1: Immediate Verification

1. User requests `employeeId` change.
2. App calls Graph `PATCH`.
3. App immediately calls Graph `GET` to verify the new value.
4. If the new value is visible, show success to the user.
5. If not, show that the request was accepted but propagation is still in progress.

This is the simplest option and may be enough if immediate read-back usually succeeds in practice.

### Option 2: Accepted Then Verified

1. User requests `employeeId` change.
2. App calls Graph `PATCH`.
3. If `PATCH` succeeds, show: "Change accepted. Propagation may take a while."
4. Start background verification.
5. Background job polls Graph until `employeeId` matches the requested value.
6. When verification succeeds, notify the user.

This option fits better if the UX should reflect distributed-system reality instead of assuming instant visibility.

## Suggested Verification Strategy

- First, do one immediate verification read after `PATCH`.
- If the value matches, return success immediately.
- If the value does not match, switch to delayed verification.
- Delayed verification can poll for a short bounded window first, for example a few seconds.
- If the value still does not appear, move the check to a background process and notify the user later.

This gives the best of both approaches:

- instant success when Graph read-back is fast
- graceful handling when propagation is delayed

## SignalR Option

If the final UX needs live feedback, the flow can be:

1. User requests the change.
2. UI shows that the request was accepted.
3. Background verification keeps checking Graph.
4. SignalR pushes a message when the new `employeeId` is verified.

This is a good fit when the user should not keep refreshing the page manually.

## Recommended State Model

Use separate states in the application:

- `Requested`
- `Accepted`
- `Verified`
- `Failed`

This is better than a simple success/failure model because it reflects the difference between:

- Graph accepting the write
- Graph read-back confirming the new value
- downstream systems eventually reflecting the change

## Current Observation

- In local testing, the immediate follow-up read returned the updated `employeeId` almost instantly.
- That suggests Graph-level verification may be fast enough for the common case.
- Because this behavior is not something we should assume for every request, it is still safer to design for possible delay.

## Recommendation For This PoC

- Keep the current `PATCH` behavior.
- Add immediate read-back verification.
- If verification succeeds immediately, return success to the user.
- If verification does not succeed immediately, return an "accepted with delay" message and continue verification in the background.
- If needed later, add SignalR for user notification once verification completes.
