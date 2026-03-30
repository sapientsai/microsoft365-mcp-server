import { Brand } from "functype/branded"

// Auth brands
export type TenantId = Brand<"TenantId", string>
export const TenantId = (value: string): TenantId => Brand("TenantId", value)

export type ClientId = Brand<"ClientId", string>
export const ClientId = (value: string): ClientId => Brand("ClientId", value)

export type ClientSecret = Brand<"ClientSecret", string>
export const ClientSecret = (value: string): ClientSecret => Brand("ClientSecret", value)

export type AccessToken = Brand<"AccessToken", string>
export const AccessToken = (value: string): AccessToken => Brand("AccessToken", value)

// Entity ID brands
export type UserId = Brand<"UserId", string>
export const UserId = (value: string): UserId => Brand("UserId", value)

export type MessageId = Brand<"MessageId", string>
export const MessageId = (value: string): MessageId => Brand("MessageId", value)

export type EventId = Brand<"EventId", string>
export const EventId = (value: string): EventId => Brand("EventId", value)

export type ContactId = Brand<"ContactId", string>
export const ContactId = (value: string): ContactId => Brand("ContactId", value)

export type DriveItemId = Brand<"DriveItemId", string>
export const DriveItemId = (value: string): DriveItemId => Brand("DriveItemId", value)

export type TeamId = Brand<"TeamId", string>
export const TeamId = (value: string): TeamId => Brand("TeamId", value)

export type ChannelId = Brand<"ChannelId", string>
export const ChannelId = (value: string): ChannelId => Brand("ChannelId", value)

export type GroupId = Brand<"GroupId", string>
export const GroupId = (value: string): GroupId => Brand("GroupId", value)

export type PlanId = Brand<"PlanId", string>
export const PlanId = (value: string): PlanId => Brand("PlanId", value)

export type PlannerTaskId = Brand<"PlannerTaskId", string>
export const PlannerTaskId = (value: string): PlannerTaskId => Brand("PlannerTaskId", value)

export type NotebookId = Brand<"NotebookId", string>
export const NotebookId = (value: string): NotebookId => Brand("NotebookId", value)

export type SectionId = Brand<"SectionId", string>
export const SectionId = (value: string): SectionId => Brand("SectionId", value)

export type PageId = Brand<"PageId", string>
export const PageId = (value: string): PageId => Brand("PageId", value)

export type TodoListId = Brand<"TodoListId", string>
export const TodoListId = (value: string): TodoListId => Brand("TodoListId", value)

export type TodoTaskId = Brand<"TodoTaskId", string>
export const TodoTaskId = (value: string): TodoTaskId => Brand("TodoTaskId", value)
