// The full tool surface, assembled from the per-domain definition modules.
//
// Order matters only for presentation: it sets the order tools are registered
// with FastMCP and therefore the order a client lists them in.

import type { ToolDefinition } from "../tool-definitions"
import { authTools } from "./auth"
import { calendarTools } from "./calendar"
import { chatsTools } from "./chats"
import { contactsTools } from "./contacts"
import { filesTools } from "./files"
import { mailTools } from "./mail"
import { meetingsTools } from "./meetings"
import { onenoteTools } from "./onenote"
import { plannerTools } from "./planner"
import { queryTools } from "./query"
import { teamsTools } from "./teams"
import { todoTools } from "./todo"
import { usersGroupsTools } from "./users-groups"

export const toolDefinitions: ReadonlyArray<ToolDefinition> = [
  ...authTools,
  ...mailTools,
  ...calendarTools,
  ...contactsTools,
  ...filesTools,
  ...chatsTools,
  ...teamsTools,
  ...meetingsTools,
  ...usersGroupsTools,
  ...plannerTools,
  ...onenoteTools,
  ...todoTools,
  ...queryTools,
]
