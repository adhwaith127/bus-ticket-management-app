# Android App — Login & Session Integration Guide

## Overview

Authentication uses **JWT tokens stored in HTTP-only cookies**.
The server manages the cookies — your app does not read or set them manually.
All requests must include cookies (`CookieJar` enabled on your HTTP client).

---

## Step 1 — Login

**POST** `/auth/login/`

**Request body:**
```json
{
  "username": "john",
  "password": "secret123",
  "device_uid": "ANDROID-UNIQUE-DEVICE-ID"
}
```

`device_uid` must be a **stable, unique string per physical device**.
Recommended: use `Settings.Secure.ANDROID_ID`, or generate a UUID once on first launch and store it in SharedPreferences. Always send it on every login — this is how the server identifies your device.

**Possible responses:**

| HTTP | `error_code` | What to do |
|---|---|---|
| 200 | — | Login successful. Cookies are set. Proceed to app. |
| 403 | `DEVICE_PENDING_APPROVAL` | First login on this device. Show "Awaiting admin approval" screen. User must wait for admin to approve the device. |
| 403 | `DEVICE_INACTIVE` | Device was revoked by admin. Show message, no retry. |
| 403 | `DEVICE_LIMIT_REACHED` | Company's device limit is full. Another device must log out first. Show appropriate message. |
| 403 | `DEVICE_UID_ALREADY_BOUND` | This device_uid is already linked to a different user account. |
| 401 | — | Wrong username or password. |
| 403 | — | Account inactive or company license not approved. Show message from server. |

**Success response body:**
```json
{
  "message": "Login Successful",
  "user": {
    "id": 5,
    "username": "john",
    "email": "john@example.com",
    "role": "user",
    "is_verified": true,
    "company_name": "ABC Travels",
    "valid_till": "31-03-2026",
    "license_status": "Approve"
  }
}
```

---

## Step 2 — Making API Calls

Just make normal requests. Cookies are sent automatically by the HTTP client (as long as `CookieJar` is enabled).

If you receive **HTTP 401** on any API call, the access token has expired. Handle it like this:

1. Call the token refresh endpoint (see Step 3)
2. If refresh succeeds → retry the original failed request
3. If refresh also fails → clear local session and take the user to the login screen

---

## Step 3 — Token Refresh

**POST** `/token/refresh/`

**Request body:**
```json
{
  "device_uid": "ANDROID-UNIQUE-DEVICE-ID"
}
```

Always include `device_uid` here. This tells the server the device is still active and resets its 24-hour inactivity timer. If you skip it, an idle device may lose its slot after 24 hours.

**Responses:**

| HTTP | Meaning |
|---|---|
| 200 | New access token issued (cookie updated automatically). Retry your original request. |
| 401 | Refresh token has expired (idle for more than 7 days). Clear session and send user to login. |

> **Note:** The refresh token rotates on every refresh — each successful refresh issues a new refresh token and invalidates the old one. This is all handled automatically via cookies. You do not need to read, store, or track tokens manually.

---

## Step 4 — Logout

**POST** `/auth/logout/`

**Request body:**
```json
{
  "device_uid": "ANDROID-UNIQUE-DEVICE-ID"
}
```

Always send `device_uid` on logout. This immediately frees the company's device slot so another device can log in. If you skip it or don't call logout (e.g. app is force-killed), the slot stays occupied for up to 24 hours before it auto-expires.

After a successful logout response, clear your local user state (SharedPreferences, etc.) and redirect to the login screen.

---

## Session Lifetime

| Token | Lifetime | What happens on expiry |
|---|---|---|
| Access token | 30 minutes | Server returns 401 → call refresh → retry original request |
| Refresh token | 7 days **sliding** | Every token refresh resets the 7-day window. Only expires if app is completely idle for 7 days. |

As long as the user opens the app at least once every 7 days, they stay logged in. If the app is truly idle for 7 days, the user must log in again.

---

## Device Slot Rules

- The company is allocated a fixed number of simultaneous mobile logins.
- Each approved device occupies one slot when active (logged in).
- The slot is freed when the device calls the logout endpoint.
- If the app crashes or is force-killed without logging out, the slot auto-frees after **24 hours** of no token refresh activity.
- Re-opening the app and refreshing the token on an already-active device does **not** consume an extra slot.

---

## HTTP Client Setup Checklist

- [ ] Enable persistent `CookieJar` so cookies are stored and sent automatically on every request.
- [ ] Point all API calls to the same base URL so cookies are scoped correctly.
- [ ] Implement a global 401 interceptor: refresh token → retry → if still 401 → go to login.
- [ ] Do **not** try to read or manually set the `access_token` or `refresh_token` cookies — they are `HttpOnly` and managed entirely by the server.

---

## Developer Summary Checklist

- [ ] Generate a stable `device_uid` once on first launch, store in SharedPreferences, reuse on every subsequent call
- [ ] Send `device_uid` on every **login**, **token refresh**, and **logout** request
- [ ] Enable `CookieJar` on your HTTP client (OkHttp, Retrofit, etc.)
- [ ] On any API 401 → call `/token/refresh/` with `device_uid` → retry once → if still failing → go to login screen
- [ ] On logout, always POST to `/auth/logout/` with `device_uid` before clearing local state
- [ ] Handle all `error_code` values from the login response with appropriate UI messages
