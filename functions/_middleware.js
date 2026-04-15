export async function onRequest(context) {
  const host = new URL(context.request.url).hostname;
  if (host !== "dm.mellender.io") {
    return new Response("Not found", { status: 404 });
  }
  return context.next();
}
