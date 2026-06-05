export default {
  async fetch(request, env) {
    try {
      const allowedOrigins = (env.ALLOWED_ORIGINS || "")
        .split(",")
        .map((x) => x.trim())
        .filter(Boolean);

      const requestOrigin = request.headers.get("Origin") || "";
      const url = new URL(request.url);
      
      // For /trigger endpoint, allow all origins (it's safe since it just dispatches CI)
      const isTriggerEndpoint = url.pathname === "/trigger";
      const allowOriginHeader =
        isTriggerEndpoint
          ? "*"
          : allowedOrigins.length === 0
            ? "*"
            : allowedOrigins.includes(requestOrigin)
              ? requestOrigin
              : "";

      const corsHeaders = {
        "Access-Control-Allow-Origin": allowOriginHeader || "*",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, X-Trigger-Key",
        "Access-Control-Max-Age": "86400",
        "Vary": "Origin",
      };

      if (request.method === "OPTIONS") {
        return new Response(null, { status: 204, headers: corsHeaders });
      }

      if (url.pathname === "/") {
        return jsonResponse({ ok: true, service: "daily-lt-funds-trigger" }, 200, corsHeaders);
      }

      if (!isTriggerEndpoint && allowedOrigins.length > 0 && requestOrigin && !allowedOrigins.includes(requestOrigin)) {
        return jsonResponse({ ok: false, error: "origin_not_allowed" }, 403, corsHeaders);
      }

      const githubToken = env.GITHUB_WORKFLOW_TOKEN;
      const repository = env.GITHUB_REPOSITORY;
      const workflowFile = env.GITHUB_WORKFLOW_FILE || "daily_publish.yml";
      const workflowRef = env.GITHUB_WORKFLOW_REF || "main";

      if (!githubToken || !repository) {
        return jsonResponse({ ok: false, error: "worker_not_configured" }, 500, corsHeaders);
      }

      if (url.pathname === "/trigger" && request.method === "POST") {
        const publicKey = env.PUBLIC_TRIGGER_KEY || "";
        if (publicKey) {
          const providedKey =
            request.headers.get("X-Trigger-Key") || url.searchParams.get("key") || "";
          if (providedKey !== publicKey) {
            return jsonResponse({ ok: false, error: "invalid_trigger_key" }, 401, corsHeaders);
          }
        }

        const dispatchedAt = new Date().toISOString();
        const ghResp = await fetch(
          `https://api.github.com/repos/${repository}/actions/workflows/${workflowFile}/dispatches`,
          {
            method: "POST",
            headers: {
              "Accept": "application/vnd.github+json",
              "Authorization": `Bearer ${githubToken}`,
              "X-GitHub-Api-Version": "2022-11-28",
              "Content-Type": "application/json",
              "User-Agent": "daily-lt-funds-trigger-worker",
            },
            body: JSON.stringify({ ref: workflowRef }),
          }
        );

        if (ghResp.status === 204) {
          return jsonResponse({ ok: true, dispatched: true, dispatchedAt }, 200, corsHeaders);
        }

        const errorText = await ghResp.text();
        return jsonResponse(
          {
            ok: false,
            error: "github_dispatch_failed",
            status: ghResp.status,
            detail: errorText.slice(0, 500),
          },
          502,
          corsHeaders
        );
      }

      if (url.pathname === "/status" && request.method === "GET") {
        const afterIso = url.searchParams.get("after") || "";
        const afterTs = Date.parse(afterIso);
        const runsResp = await fetch(
          `https://api.github.com/repos/${repository}/actions/workflows/${workflowFile}/runs?event=workflow_dispatch&branch=${encodeURIComponent(workflowRef)}&per_page=20`,
          {
            headers: {
              "Accept": "application/vnd.github+json",
              "Authorization": `Bearer ${githubToken}`,
              "X-GitHub-Api-Version": "2022-11-28",
              "User-Agent": "daily-lt-funds-trigger-worker",
            },
          }
        );

        if (!runsResp.ok) {
          const errorText = await runsResp.text();
          return jsonResponse(
            {
              ok: false,
              error: "github_runs_failed",
              status: runsResp.status,
              detail: errorText.slice(0, 500),
            },
            502,
            corsHeaders
          );
        }

        const runsPayload = await runsResp.json();
        const runs = (runsPayload.workflow_runs || []).filter((run) => {
          if (!afterIso || Number.isNaN(afterTs)) {
            return true;
          }
          const createdTs = Date.parse(run.created_at || "");
          return !Number.isNaN(createdTs) && createdTs >= afterTs;
        });

        const run = runs.length > 0 ? runs[0] : null;
        return jsonResponse(
          {
            ok: true,
            run: run
              ? {
                  id: run.id,
                  status: run.status,
                  conclusion: run.conclusion,
                  createdAt: run.created_at,
                  updatedAt: run.updated_at,
                  htmlUrl: run.html_url,
                }
              : null,
          },
          200,
          corsHeaders
        );
      }

      return jsonResponse({ ok: false, error: "not_found" }, 404, corsHeaders);
    } catch (error) {
      const corsHeaders = {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, X-Trigger-Key",
      };
      return jsonResponse(
        {
          ok: false,
          error: "worker_error",
          detail: error && error.message ? error.message : "Internal server error",
        },
        500,
        corsHeaders
      );
    }
  },
};

function jsonResponse(payload, status, headers) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      ...headers,
    },
  });
}
