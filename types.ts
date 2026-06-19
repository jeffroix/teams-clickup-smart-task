export interface TaskProposal {
  title: string;
  description: string;
  assignee_hint: string | null;
  due_date: string | null;
  priority: "urgent" | "high" | "normal" | "low" | null;
  source_summary: string;
  open_questions: string[];
}

export interface TeamsMessageContext {
  messageId: string | null;
  selectedText: string;
  senderName: string | null;
  senderId: string | null;
  conversationId: string | null;
  chatId: string | null;
  teamId: string | null;
  channelId: string | null;
  serviceUrl: string | null;
  tenantId: string | null;
  messageLink: string | null;
  threadMessages: string[];
  rawContext: Record<string, unknown>;
}

export interface ClickUpTaskResult {
  id: string;
  name: string;
  url: string;
}

export interface AppConfig {
  openAiApiKey: string;
  openAiModel: string;
  clickupApiToken: string;
  clickupListId: string;
  clickupAssigneeMap: Record<string, number>;
  microsoftAppId: string;
  microsoftAppPassword: string;
  microsoftTenantId: string | null;
  teamsThreadFetchEnabled: boolean;
  port: number;
}

export interface Logger {
  info(event: string, data?: Record<string, unknown>): void;
  warn(event: string, data?: Record<string, unknown>): void;
  error(event: string, data?: Record<string, unknown>): void;
}

export interface FetchTaskPreviewData {
  proposal: TaskProposal;
  context: TeamsMessageContext;
}

export interface MessageExtensionDependencies {
  extractor: {
    extractTaskProposal(context: TeamsMessageContext): Promise<TaskProposal>;
  };
  clickup: {
    createTaskFromProposal(proposal: TaskProposal, context: TeamsMessageContext): Promise<ClickUpTaskResult>;
  };
  logger: Logger;
  appId: string;
  appPassword: string;
  tenantId: string | null;
  threadFetchEnabled: boolean;
}
