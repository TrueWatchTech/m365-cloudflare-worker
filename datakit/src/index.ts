import { DurableObject } from "cloudflare:workers";

interface Env {
  CONTAINER: DurableObjectNamespace<Container>;
}

export class Container extends DurableObject<Env> {
  private readonly containerRef: NonNullable<DurableObjectState["container"]>;

  constructor(ctx: DurableObjectState, env: Env) {
    super(ctx, env);
    if (!ctx.container) {
      throw new Error("Container runtime is unavailable");
    }
    this.containerRef = ctx.container;

    void this.ctx.blockConcurrencyWhile(async () => {
      if (!this.containerRef.running) {
        this.containerRef.start({
          enableInternet: true,
          env: {
            // Your TrueWatch DataWay URL with token.
            // Format: https://openway.truewatch.com?token=YOUR_TOKEN
            // Find this in your TrueWatch workspace settings.
            ENV_DATAWAY: "YOUR_TRUEWATCH_DATAWAY_URL",
            ENV_GLOBAL_HOST_TAGS: "datakit-source=cloudflare",
            ENV_HTTP_LISTEN: "0.0.0.0:9529",
            ENV_DEFAULT_ENABLED_INPUTS: "logstreaming",
            ENV_HTTP_PUBLIC_APIS: "/, /metrics, /logstreaming",
          },
        });
      }
    });
  }

  async fetch(request: Request): Promise<Response> {
    try {
      const body =
        request.method === "GET" || request.method === "HEAD"
          ? undefined
          : await request.arrayBuffer();

      const upstreamUrl = request.url.replace("https:", "http:");
      const upstreamRequest = new Request(upstreamUrl, {
        method: request.method,
        headers: request.headers,
        body,
        redirect: "manual",
      });

      return this.containerRef.getTcpPort(9529).fetch(upstreamRequest);
    } catch (error) {
      const message = formatError(error);
      console.error("Error fetch:", message);

      return new Response(
        JSON.stringify({
          id: this.ctx.id.toString(),
          error: message,
          timestamp: Date.now(),
          status: 500,
        }),
        {
          status: 500,
          headers: { "Content-Type": "application/json" },
        },
      );
    }
  }
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    try {
      return env.CONTAINER.get(env.CONTAINER.idFromName("fetcher")).fetch(request);
    } catch (error) {
      const message = formatError(error);
      console.error("Error fetch:", message);

      return new Response(
        JSON.stringify({
          error: message,
          timestamp: Date.now(),
        }),
        {
          status: 500,
          headers: { "Content-Type": "application/json" },
        },
      );
    }
  },
};

function formatError(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  if (typeof error === "object" && error !== null && "message" in error) {
    return String((error as { message: unknown }).message);
  }
  return String(error);
}
