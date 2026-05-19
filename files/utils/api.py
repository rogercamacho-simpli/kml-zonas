import requests

API_BASE = "http://api.simpliroute.com"
API_GW   = "https://api-gateway.simpliroute.com"
API_V2   = "https://api.simpliroute.com"


def sr_request(method, endpoint, token, base=API_BASE, extra_headers=None, **kwargs):
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    if extra_headers:
        headers.update(extra_headers)
    try:
        r = requests.request(
            method,
            f"{base}{endpoint}",
            headers=headers,
            timeout=kwargs.pop("timeout", 30),
            **kwargs,
        )
        return r.status_code, r.json() if r.content else {}
    except Exception as e:
        return None, str(e)


# ── Shortcuts ─────────────────────────────────────────────────────────────────
def sr_get(endpoint, token, base=API_BASE, **kwargs):
    return sr_request("GET", endpoint, token, base=base, **kwargs)

def sr_post(endpoint, token, payload, base=API_BASE, extra_headers=None, **kwargs):
    return sr_request("POST", endpoint, token, base=base, extra_headers=extra_headers, json=payload, **kwargs)

def sr_put(endpoint, token, payload, base=API_BASE, **kwargs):
    return sr_request("PUT", endpoint, token, base=base, json=payload, **kwargs)

def sr_patch(endpoint, token, payload, base=API_BASE, **kwargs):
    return sr_request("PATCH", endpoint, token, base=base, json=payload, **kwargs)

def sr_delete(endpoint, token, base=API_BASE, **kwargs):
    return sr_request("DELETE", endpoint, token, base=base, **kwargs)


# ── Helpers de usuario ────────────────────────────────────────────────────────
def get_users_list(token):
    return sr_get("/v1/accounts/users/", token, timeout=300)


def put_user_full(token, user, new_email):
    payload = {
        "id":               user["id"],
        "username":         user.get("username", ""),
        "name":             user.get("name", ""),
        "phone":            user.get("phone", ""),
        "email":            new_email,
        "is_owner":         user.get("is_owner", False),
        "is_admin":         user.get("is_admin", False),
        "is_driver":        user.get("is_driver", False),
        "is_codriver":      user.get("is_codriver", False),
        "is_router_jr":     user.get("is_router_jr", False),
        "is_monitor":       user.get("is_monitor", False),
        "is_coordinator":   user.get("is_coordinator", False),
        "is_router":        user.get("is_router", False),
        "is_staff":         user.get("is_staff", False),
        "is_seller_viewer": user.get("is_seller_viewer", False),
        "is_seller":        user.get("is_seller", False),
        "blocked":          user.get("blocked", False),
        "status":           user.get("status", "active"),
    }
    return sr_put(f"/v1/accounts/users/{user['id']}/", token, payload)