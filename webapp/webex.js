import Http from "./http.js";

class Webex {
  #http;

  /**
   * Creates instance of Webex API Integration
   * @param {string} accessToken - Required auth
   * @param {?URL} baseUrl
   */
  constructor(accessToken, baseUrl = "https://webexapis.com/v1") {
    if (
      !accessToken ||
      typeof accessToken !== "string" ||
      !accessToken.trim()
    ) {
      throw new Error("Access token is required");
    }

    if (!baseUrl || typeof baseUrl !== "string" || !baseUrl.trim()) {
      throw new Error("Base URL is required");
    }

    this.#http = new Http(accessToken, baseUrl);
  }

  /** @typedef {object} jsonWebhooks
   * @property {object[]} items
   * @property {string} items.id
   * @property {string} items.name
   * @property {string} items.targetUrl
   * @property {string} items.resource
   * @property {string} items.event
   * @property {?string} items.filter
   * @property {string} items.orgId
   * @property {string} items.createdBy
   * @property {string} items.appId
   * @property {string} items.ownedBy
   * @property {string} items.status
   * @property {string} items.created
   */

  /**
   * List all of your webhooks.
   * @param {?object} params - Optional: params for the request
   * @param {?number} params.max - Optional: Limit the maximum number of webhooks in the response. example = "100", default = 100
   * @param {?string} params.ownedBy - Optional: Limit the result list to org wide webhooks. Only allowed value is org.example = "org"
   * @returns {Promise<jsonWebhooks>} A promise that resolves with the webhooks information or rejects with an error.
   * @see {@link https://developer.webex.com/docs/api/v1/webhooks/list-webhooks Webex API Documentation}
   */
  async listWebhooks(params = {}, onProgress = null, onComplete = null) {
    console.log("Listing Webhooks");
    return this.#http.getPaginated("/webhooks", params, {
      onProgress,
      onComplete,
    });
  }

  /**
   * Update a webhook
   * @param {string} webhookId - The unique identifier for the webhook. example = "Y2lzY29zcGFyazovL3VzL1dFQkhPT0svOTZhYmMyYWEtM2RjYy0xMWU1LWExNTItZmUzNDgxOWNkYzlh"
   * @param {object} params - Params for the request
   * @param {string} params.name - A user-friendly name for the webhook.
   * @param {string} params.targetUrl - URL that receives POST requests for each event.
   * @param {?string} params.secret - Optional: Secret used to generate payload signature.
   * @param {?string} params.ownedBy - Optional: Specify org when creating an org/admin level webhook.
   * @param {?string} params.status - Optional: Status of the webhook. Use "active" to reactivate a disabled webhook.
   */
  async updateWebhook(webhookId, params = {}) {
    console.log("Updating Webhook Id:", webhookId);
    return this.#http.put("/webhooks/" + webhookId, params);
  }

  /**
   * Deletes a webhook, by ID.
   * @param {string} webhookId - The unique identifier for the webhook. example = "Y2lzY29zcGFyazovL3VzL1dFQkhPT0svOTZhYmMyYWEtM2RjYy0xMWU1LWExNTItZmUzNDgxOWNkYzlh"
   */
  async deleteWebhook(webhookId) {
    console.log("Deleting Webhook Id:", webhookId);
    return this.#http.delete("/webhooks/" + webhookId);
  }

  /**
   * Create a Webhook
   * @param {object} params - Optional: params for the request
   * @param {string} params.name - A user-friendly name for the webhook.
   * @param {string} params.targetUrl - URL that receives POST requests for each event.
   * @param {string} params.resource - Resource type for the webhook - ["attachmentActions", "dataSources", "memberships", "messages", "rooms", ...]
   * @param {string} params.event - Event type for the webhook - ["created", "updated", "deleted", "started", "ended", ...]
   * @param {string} params.filter - Optional: Filter that defines the webhook scope.
   * @param {?string} params.secret - Optional: Secret used to generate payload signature.
   * @param {?string} params.ownedBy - Optional: Specify org when creating an org/admin level webhook.
   */
  async createWebhook(params = {}) {
    console.log("Creating Webhook");
    return this.#http.post("/webhooks", params);
  }

  /** @typedef {object} jsonOrgDetails
   * @property {object[]} item
   * @property {string} item.id
   * @property {string} item.displayName
   * @property {string} item.created
   */

  /**
   * Shows details for an organization, by ID.
   * @param {string} orgId - The unique identifier for the message. example = "Y2lzY29zcGFyazovL3VzL01FU1NBR0UvOTJkYjNiZTAtNDNiZC0xMWU2LThhZTktZGQ1YjNkZmM1NjVk"
   * @returns {Promise<jsonOrgDetails>} A promise that resolves with the organization details information or rejects with an error.
   * @see {@link https://developer.webex.com/admin/docs/api/v1/organizations/get-organization-details Webex API Documentation}
   */
  async getOrgDetails(orgId) {
    console.log("Getting org details");
    return this.#http.getJson("/organizations/" + orgId);
  }

  /**
   * Post a plain text or rich text message, and optionally, a file attachment attachment, to a room.
   * @param {object} params - Optional: params for the request
   * @param {?string} params.roomId - Optional: The room ID of the message.
   * @param {?string} params.parentId - Optional: The parent message to reply to.
   * @param {?string} params.toPersonId - Optional: The person ID of the recipient when sending a private 1:1 message.
   * @param {?string} params.toPersonEmail - Optional: The email address of the recipient when sending a private 1:1 message.
   * @param {?string} params.text - Optional: The message, in plain text. If markdown is specified this parameter may be optionally used to provide alternate text for UI clients that do not support rich text. The maximum message length is 7439 bytes.
   * @param {?string} params.markdown - Optional: The message, in Markdown format. The maximum message length is 7439 bytes.
   * @param {?string[]} params.files - Optional: The public URL to a binary file to be posted into the room. Only one file is allowed per message. Uploaded files are automatically converted into a format that all Webex clients can render.
   * @param {?object[]} params.attachments - Optional: Content attachments to attach to the message. Only one card per message is supported. See the Cards Guide for more information.
   */
  async createMessage(params = {}) {
    console.log("Creating Message");
    return this.#http.post("/messages", params);
  }

  /**
   * Show details for a message, by message ID.
   * @param {string} messageId - The unique identifier for the message. example = "Y2lzY29zcGFyazovL3VzL01FU1NBR0UvOTJkYjNiZTAtNDNiZC0xMWU2LThhZTktZGQ1YjNkZmM1NjVk"
   */
  async getMessage(messageId) {
    console.log("Creating Message");
    return this.#http.getJson("/messages/" + messageId);
  }

  /**
   * Update a message you have posted not more than 10 times.
   * @param {string} messageId - Optional: params for the request
   * @param {object} params - Optional: params for the request
   * @param {?string} params.text - Optional: The message, in plain text. If markdown is specified this parameter may be optionally used to provide alternate text for UI clients that do not support rich text. The maximum message length is 7439 bytes.
   * @param {?string} params.markdown - Optional: The message, in Markdown format. The maximum message length is 7439 bytes.
   * @param {?object[]} params.attachments - Optional: Content attachments to attach to the message. Only one card per message is supported. See the Cards Guide for more information.
   */
  async editMessage(messageId, params = {}) {
    console.log("Creating Message");
    return this.#http.put("/messages/" + messageId, params);
  }

  /**
   * Delete a message, by message ID
   * @param {string} messageId - The unique identifier for the message. example = "Y2lzY29zcGFyazovL3VzL01FU1NBR0UvOTJkYjNiZTAtNDNiZC0xMWU2LThhZTktZGQ1YjNkZmM1NjVk"
   */
  async deleteMessage(messageId, params = {}) {
    console.log("Creating Message");
    return this.#http.delete("/messages/" + messageId);
  }

  /** @typedef {object} jsonPersonDetails
   * @property {string[]} emails
   * @property {object[]} phoneNumbers
   * @property {string} phoneNumbers.type
   * @property {string} phoneNumbers.value
   * @property {boolean} phoneNumbers.primary
   * @property {string} extension
   * @property {string} locationId
   * @property {string} displayName
   * @property {string} nickName
   * @property {string} firstName
   * @property {string} lastName
   * @property {string} avatar
   * @property {string} orgId
   * @property {string[]} roles
   * @property {string[]} licenses
   * @property {string} department
   * @property {string} manager
   * @property {string} managerId
   * @property {string} title
   * @property {object[]} addresses
   * @property {string} addresses.type
   * @property {string} addresses.country
   * @property {string} addresses.locality
   * @property {string} addresses.region
   * @property {string} addresses.streetAddress
   * @property {string} addresses.postalCode
   * @property {string} created
   * @property {string} lastModified
   * @property {string} timezone
   * @property {string} lastActivity
   * @property {string[]} siteUrls
   * @property {object[]} sipAddresses
   * @property {string} sipAddresses.type
   * @property {string} sipAddresses.value
   * @property {boolean} sipAddresses.primary
   * @property {string} xmppFederationJid
   * @property {string} status
   * @property {string} invitePending
   * @property {string} loginEnabled
   * @property {string} type
   */

  /**
   * Get profile details for the authenticated user.
   * @param {string} messageId - The unique identifier for the message. example = "Y2lzY29zcGFyazovL3VzL01FU1NBR0UvOTJkYjNiZTAtNDNiZC0xMWU2LThhZTktZGQ1YjNkZmM1NjVk"
   * @returns {Promise<jsonPersonDetails>} A promise that resolves with the person details information or rejects with an error.
   * @see {@link https://developer.webex.com/admin/docs/api/v1/people/get-my-own-details Webex API Documentation}
   */
  async getMyOwnDetails() {
    console.log("Getting my own details");
    return this.#http.getJson("/people/me");
  }

  /**
   * Shows details for a person, by ID.
   * @param {string} personId - A unique identifier for the person. example = "Y2lzY29zcGFyazovL3VzL1BFT1BMRS9mNWIzNjE4Ny1jOGRkLTQ3MjctOGIyZi1mOWM0NDdmMjkwNDY"
   * @param {object} params - Optional: params for the request
   * @param {?boolean} params.callingData - Optional: Include Webex Calling user details in the response. true or false
   * @returns {Promise<jsonPersonDetails>} A promise that resolves with the person details information or rejects with an error.
   * @see {@link https://developer.webex.com/admin/docs/api/v1/people/get-person-details Webex API Documentation}
   */
  async getPersonDetails(personId) {
    console.log("Getting Person Details");
    return this.#http.getJson("/people/" + personId);
  }

  /**
   * List people in your organization. For most users, either the email or displayName parameter is required.
   * @param {?object} params - Optional: params for the request
   * @param {?string} params.email - Optional: List people with this email address
   * @param {?string} params.displayName - Optional: List people whose name starts with this string.
   * @param {?string} params.id - Optional: List people by ID. Accepts up to 85 person IDs separated by commas.
   * @param {?number} params.max - Optional: Limit the maximum number of people in the response
   * @returns {Promise<jsonPersonDetails[]>} A promise that resolves with the people information or rejects with an error.
   * @see {@link https://developer.webex.com/messaging/docs/api/v1/people/list-people Webex API Documentation}
   */
  async listPeople(params) {
    console.log("List People");
    return this.#http.getPaginated("/people", params);
  }

  /**
   * Shows details for a attachment action, by ID.
   * @param {string} actionId - A unique identifier for the attachment action.
   */
  async getActionDetails(actionId) {
    console.log("Creating Message");
    return this.#http.getJson("/attachment/actions/" + actionId);
  }

  /** @typedef {object} jsonRoomDetails
   * @property {string} id
   * @property {string} title
   * @property {string} type
   * @property {boolean} isLocked
   * @property {string} teamId
   * @property {string} lastActivity
   * @property {string} creatorId
   * @property {string} created
   * @property {string} ownerId
   * @property {string} classificationId
   * @property {boolean} isAnnouncementOnly
   * @property {boolean} isReadOnly
   * @property {boolean} isPublic
   * @property {string} madePublic
   * @property {string} description
   */

  /**
   * Shows details for a room, by ID.
   * @param {string} roomId -The unique identifier for the room. example = "Y2lzY29zcGFyazovL3VzL1BFT1BMRS9mNWIzNjE4Ny1jOGRkLTQ3MjctOGIyZi1mOWM0NDdmMjkwNDY"
   * @returns {Promise<jsonRoomDetails>} A promise that resolves with the room details information or rejects with an error.
   * @see {@link https://developer.webex.com/messaging/docs/api/v1/rooms/get-room-details Webex API Documentation}
   */
  async getRoomDetails(roomId) {
    console.log("Getting Room Details");
    return this.#http.getJson("/rooms/" + roomId);
  }

  /**
   * List rooms to which the authenticated user belongs to.
   * @param {?object} params - Optional: params for the request
   * @param {?string} params.teamId - Optional: List rooms associated with a team, by ID. Cannot be set in combination with orgPublicSpaces
   * @param {?{'direct'|'group'}} params.type - Optional: List rooms by type. Cannot be set in combination with orgPublicSpaces
   * @param {?string} params.orgPublicSpaces - Optional: Shows the org's public spaces joined and unjoined. When set the result list is sorted by the madePublic timestamp.
   * @param {?string} params.from - Optional: Filters rooms, that were made public after this time. See madePublic timestamp
   * @param {?string} params.to - Optional: Filters rooms, that were made public before this time. See maePublic timestamp
   * @param {?string} params.sortBy - Optional: Sort results. Cannot be set in combination with orgPublicSpaces
   * @param {?string} params.max - Optional: Limit the maximum number of rooms in the response. Value must be between 1 and 1000, inclusive. Default is 100.
   * @returns {Promise<jsonRoomDetails[]>} A promise that resolves with the rooms information or rejects with an error.
   * @see {@link https://developer.webex.com/messaging/docs/api/v1/rooms/list-rooms Webex API Documentation}
   */
  async listRooms(params = {}, onProgress = null, onComplete = null) {
    console.log("Listing Webhooks");
    return this.#http.getPaginated("/rooms", params, {
      onProgress,
      onComplete,
    });
  }
}

export default Webex;
