import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  findByUserId: vi.fn(),
  get: vi.fn(),
  list: vi.fn(),
  upsert: vi.fn(),
  createOneOnOneChatViaGraph: vi.fn(),
  resolveGraphToken: vi.fn(),
  resolveMSTeamsCredentials: vi.fn(),
  loadMSTeamsSdkWithAuth: vi.fn(),
}));

vi.mock("./conversation-store-fs.js", () => ({
  createMSTeamsConversationStoreFs: () => ({
    findByUserId: mocks.findByUserId,
    get: mocks.get,
    list: mocks.list,
    upsert: mocks.upsert,
  }),
}));

vi.mock("./graph.js", () => ({
  createOneOnOneChatViaGraph: mocks.createOneOnOneChatViaGraph,
  resolveGraphToken: mocks.resolveGraphToken,
}));

vi.mock("./token.js", () => ({
  resolveMSTeamsCredentials: mocks.resolveMSTeamsCredentials,
}));

vi.mock("./sdk.js", () => ({
  createMSTeamsAdapter: () => ({ continueConversation: vi.fn(), process: vi.fn() }),
  loadMSTeamsSdkWithAuth: mocks.loadMSTeamsSdkWithAuth,
}));

vi.mock("./runtime.js", () => ({
  getMSTeamsRuntime: () => ({
    logging: {
      getChildLogger: () => ({
        info: vi.fn(),
        debug: vi.fn(),
      }),
    },
  }),
}));

describe("resolveMSTeamsSendContext", () => {
  beforeEach(() => {
    mocks.get.mockReset();
    mocks.findByUserId.mockReset();
    mocks.list.mockReset();
    mocks.upsert.mockReset();
    mocks.createOneOnOneChatViaGraph.mockReset();
    mocks.resolveGraphToken.mockReset();
    mocks.resolveMSTeamsCredentials.mockReset();
    mocks.loadMSTeamsSdkWithAuth.mockReset();

    mocks.resolveMSTeamsCredentials.mockReturnValue({
      appId: "app-1",
      appPassword: "secret",
      tenantId: "tenant-1",
    });
    mocks.loadMSTeamsSdkWithAuth.mockResolvedValue({
      sdk: {
        MsalTokenProvider: class {
          getAccessToken() {
            return "token";
          }
        },
      },
      authConfig: {},
    });
  });

  it("bootstraps a personal conversation via Graph when no stored reference exists", async () => {
    mocks.findByUserId.mockResolvedValueOnce(null);
    mocks.resolveGraphToken.mockResolvedValueOnce("graph-token");
    mocks.createOneOnOneChatViaGraph.mockResolvedValueOnce("19:new-chat@thread.v2");
    mocks.list.mockResolvedValueOnce([]);

    const { resolveMSTeamsSendContext } = await import("./send-context.js");

    const result = await resolveMSTeamsSendContext({
      cfg: { channels: { msteams: { enabled: true } } },
      to: "user:user-aad-id",
    });

    expect(result.conversationId).toBe("19:new-chat@thread.v2");
    expect(result.conversationType).toBe("personal");
    expect(mocks.upsert).toHaveBeenCalledWith(
      "19:new-chat@thread.v2",
      expect.objectContaining({
        channelId: "msteams",
        serviceUrl: "https://smba.trafficmanager.net/teams/",
        conversation: expect.objectContaining({
          conversationType: "personal",
        }),
      }),
    );
  });

  it("throws actionable error when bootstrap fails", async () => {
    mocks.findByUserId.mockResolvedValueOnce(null);
    mocks.resolveGraphToken.mockRejectedValueOnce(new Error("forbidden"));

    const { resolveMSTeamsSendContext } = await import("./send-context.js");

    await expect(
      resolveMSTeamsSendContext({
        cfg: { channels: { msteams: { enabled: true } } },
        to: "user:user-aad-id",
      }),
    ).rejects.toThrow("Tried creating a new personal chat via Graph");
  });
});
