/**
 * G3AxonLive Client Engine v2.0
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Lean vanilla JavaScript engine for reactive ASP components.
 * Intercepts component events, sends async fetch requests to the server,
 * and performs targeted DOM swaps without full page reloads.
 *
 * Features:
 *   - Event interception (click, change, submit) on data-g3al-* elements
 *   - Targeted outerHTML DOM swaps via JSON patch response
 *   - Server-triggered client actions: set_timer, redirect, trigger, add_attribute
 *   - Exponential backoff retry for transient network errors
 *   - Global error handler hook
 */

(function () {
    'use strict';

    /**
     * G3AxonLive: Global client engine namespace for reactive ASP components.
     * All state and methods are encapsulated to prevent global namespace pollution.
     */
    window.G3AxonLive = {
        // Session ID transmitted with every fetch request; set via init()
        sessionId: null,

        // Tracks whether an async fetch operation is currently in flight (debounce guard)
        isProcessing: false,

        // Endpoint URL for all G3AxonLive fetch requests
        endpoint: '/g3al',

        // Retry configuration for exponential backoff on transient network errors
        maxRetries: 3,
        retryBaseDelayMs: 1000,

        /**
         * Initialize the G3AxonLive engine on page load.
         * Must be called with the session ID provided by the server-side framework.
         * @param {string} sessionId - The user's ASP Session ID
         * @returns {boolean} - True if initialization succeeded
         */
        init: function (sessionId) {
            if (!sessionId) {
                console.error('G3AxonLive: sessionId is required for initialization');
                return false;
            }
            this.sessionId = sessionId;
            this.attachComponentEventHandlers();
            return true;
        },

        /**
         * Attach delegated click, change, and submit event handlers at the document level.
         * Reactive components are identified by the data-g3al-id attribute.
         * Event type is specified by the data-g3al-event attribute (click, change, submit).
         */
        attachComponentEventHandlers: function () {
            var self = this;

            // Intercept click events on reactive components
            document.addEventListener('click', function (e) {
                var component = self.findComponentElement(e.target);
                if (component && component.getAttribute('data-g3al-event') === 'click') {
                    e.preventDefault();
                    e.stopPropagation();
                    var componentId = component.getAttribute('data-g3al-id');
                    var eventName = component.getAttribute('data-g3al-event-name') || 'onclick';
                    var eventArgs = self.extractEventArgs(component);
                    self.sendEvent(componentId, eventName, eventArgs);
                }
            }, true);

            // Intercept change events on form inputs and selects
            document.addEventListener('change', function (e) {
                var component = self.findComponentElement(e.target);
                if (component && component.getAttribute('data-g3al-event') === 'change') {
                    e.preventDefault();
                    e.stopPropagation();
                    var componentId = component.getAttribute('data-g3al-id');
                    var eventName = component.getAttribute('data-g3al-event-name') || 'onchange';
                    var eventArgs = self.extractEventArgs(component);
                    self.sendEvent(componentId, eventName, eventArgs);
                }
            }, true);

            // Intercept form submission from reactive component containers
            document.addEventListener('submit', function (e) {
                var form = e.target;
                if (form && form.getAttribute('data-g3al-component') === 'true') {
                    e.preventDefault();
                    e.stopPropagation();
                    var componentId = form.getAttribute('data-g3al-id');
                    var eventName = form.getAttribute('data-g3al-event-name') || 'onsubmit';
                    var eventArgs = self.extractEventArgs(form);
                    self.sendEvent(componentId, eventName, eventArgs);
                }
            }, true);
        },

        /**
         * Walk up the DOM tree to find the closest ancestor with data-g3al-id.
         * @param {HTMLElement} element - Starting DOM element (event target)
         * @returns {HTMLElement|null} - Nearest reactive component element, or null
         */
        findComponentElement: function (element) {
            var el = element;
            while (el && el !== document) {
                if (el.getAttribute && el.getAttribute('data-g3al-id')) {
                    return el;
                }
                el = el.parentNode;
            }
            return null;
        },

        /**
         * Extract event arguments from data-g3al-arg-* attributes on the component.
         * For example: data-g3al-arg-step="2" yields { step: "2" }.
         * @param {HTMLElement} component - The reactive component element
         * @returns {Object} - Map of argument names to string values
         */
        extractEventArgs: function (component) {
            var args = {};
            if (!component.attributes) return args;
            for (var i = 0; i < component.attributes.length; i++) {
                var attr = component.attributes[i];
                if (attr.name.indexOf('data-g3al-arg-') === 0) {
                    var argName = attr.name.substring('data-g3al-arg-'.length);
                    args[argName] = attr.value;
                }
            }
            return args;
        },

        /**
         * Send an asynchronous component event to the server with exponential backoff retry.
         * Implements retry delays of 1s, 2s, 4s for transient network errors.
         * Debounces concurrent requests; a second event while one is in flight is dropped.
         * @param {string} componentId - ID of the component firing the event
         * @param {string} eventName   - Event name (e.g. "onclick", "onchange")
         * @param {Object} eventArgs   - Optional key/value event arguments
         */
        sendEvent: function (componentId, eventName, eventArgs) {
            if (this.isProcessing) {
                console.warn('G3AxonLive: Event dropped — another request is in flight');
                return;
            }
            if (!this.sessionId) {
                console.error('G3AxonLive: Session ID not initialized — call G3AxonLive.init(sessionId) first');
                return;
            }

            var self = this;
            this.isProcessing = true;

            var payload = {
                sessionId: this.sessionId,
                componentId: componentId,
                eventName: eventName,
                eventArgs: eventArgs || {}
            };

            this._fetchWithRetry(payload, 0, function () {
                self.isProcessing = false;
            });
        },

        /**
         * Internal helper: perform a fetch POST with exponential backoff retry.
         * Retries only on network-level failures, not on HTTP 4xx/5xx application errors.
         * @param {Object}   payload  - JSON payload to POST
         * @param {number}   attempt  - Current zero-based attempt count
         * @param {Function} onDone   - Callback invoked when request resolves (success or failure)
         */
        _fetchWithRetry: function (payload, attempt, onDone) {
            var self = this;
            fetch(this.endpoint, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-G3AxonLive': 'true'
                },
                body: JSON.stringify(payload)
            })
                .then(function (response) {
                    if (!response.ok) {
                        // HTTP application error — do not retry.
                        // Try to extract JSON payload error for diagnostics.
                        return response.text().then(function (txt) {
                            var msg = response.statusText;
                            if (txt) {
                                try {
                                    var parsed = JSON.parse(txt);
                                    if (parsed && parsed.error) {
                                        msg = parsed.error;
                                    }
                                } catch (e) {
                                    // Non-JSON body; keep status text.
                                }
                            }
                            throw new Error('HTTP ' + response.status + ': ' + msg);
                        });
                    }
                    return response.json();
                })
                .then(function (data) {
                    self.processResponse(data);
                    onDone();
                })
                .catch(function (error) {
                    var isHTTPError = error.message && error.message.indexOf('HTTP') === 0;
                    var isParseError = error instanceof SyntaxError;
                    var isNetworkError = !isHTTPError && !isParseError;
                    if (isNetworkError && attempt < self.maxRetries) {
                        var delay = self.retryBaseDelayMs * Math.pow(2, attempt);
                        console.warn('G3AxonLive: Network error, retrying in ' + delay + 'ms (attempt ' + (attempt + 1) + '/' + self.maxRetries + '):', error);
                        setTimeout(function () {
                            self._fetchWithRetry(payload, attempt + 1, onDone);
                        }, delay);
                    } else {
                        console.error('G3AxonLive: Fetch error:', error);
                        self.onError(error);
                        onDone();
                    }
                });
        },

        /**
         * Process the JSON envelope returned by the server after an event.
         * First applies component HTML patches, then executes server-triggered actions.
         * @param {Object} response - Parsed JSON response from /g3al/
         */
        processResponse: function (response) {
            if (!response) {
                console.error('G3AxonLive: Empty response from server');
                return;
            }
            if (!response.success) {
                this.onError(new Error(response.error || 'Server returned success: false'));
                return;
            }

            // Apply HTML patches to update reactive component DOM nodes
            if (response.components && response.components.length > 0) {
                for (var i = 0; i < response.components.length; i++) {
                    this.applyPatch(response.components[i]);
                }
                // Re-attach event handlers after DOM nodes may have been replaced
                this.attachComponentEventHandlers();
            }

            // Execute server-triggered client actions (set_timer, redirect, trigger, add_attribute)
            if (response.actions && response.actions.length > 0) {
                this.processActions(response.actions);
            }
        },

        /**
         * Apply a single HTML patch by replacing the outerHTML of the target component.
         * The "html" property (lowercase) carries the full rendered HTML from the server.
         * @param {Object} patch - Object with componentId (string) and html (string)
         */
        applyPatch: function (patch) {
            if (!patch.componentId || patch.html === undefined) {
                console.warn('G3AxonLive: Invalid patch — expected componentId and html:', patch);
                return;
            }
            var component = document.getElementById(patch.componentId);
            if (!component) {
                console.warn('G3AxonLive: Component element not found in DOM:', patch.componentId);
                return;
            }
            try {
                component.outerHTML = patch.html;
            } catch (e) {
                console.error('G3AxonLive: Failed to apply patch for', patch.componentId, ':', e);
            }
        },

        /**
         * Execute a list of server-triggered client actions in order.
         * Each action object must have a "type" field that determines its behavior.
         *
         * Supported action types:
         *   set_timer    — schedule a component event after a delay (ms)
         *   redirect     — navigate the browser to a new URL
         *   trigger      — immediately fire a component event
         *   add_attribute — set an HTML attribute on a component element
         *
         * @param {Array} actions - Array of action objects from the server response
         */
        processActions: function (actions) {
            var self = this;
            for (var i = 0; i < actions.length; i++) {
                (function (action) {
                    switch (action.type) {
                        case 'set_timer':
                            // Schedule a server-defined event after a delay in milliseconds.
                            if (action.componentId && action.eventName && action.delay > 0) {
                                setTimeout(function () {
                                    self.sendEvent(action.componentId, action.eventName, {});
                                }, action.delay);
                            } else {
                                console.warn('G3AxonLive: set_timer missing required fields:', action);
                            }
                            break;

                        case 'redirect':
                            // Navigate the browser to the provided URL.
                            if (action.url) {
                                window.location.href = action.url;
                            } else {
                                console.warn('G3AxonLive: redirect missing url:', action);
                            }
                            break;

                        case 'trigger':
                            // Immediately fire a component event without a debounce guard.
                            if (action.componentId && action.eventName) {
                                self.sendEvent(action.componentId, action.eventName, {});
                            } else {
                                console.warn('G3AxonLive: trigger missing required fields:', action);
                            }
                            break;

                        case 'add_attribute':
                            // Set an attribute on the identified component element in the DOM.
                            if (action.componentId && action.name !== undefined) {
                                var el = document.getElementById(action.componentId);
                                if (el) {
                                    el.setAttribute(action.name, action.value || '');
                                } else {
                                    console.warn('G3AxonLive: add_attribute — element not found:', action.componentId);
                                }
                            } else {
                                console.warn('G3AxonLive: add_attribute missing required fields:', action);
                            }
                            break;

                        default:
                            console.warn('G3AxonLive: Unknown action type "' + action.type + '":', action);
                            break;
                    }
                })(actions[i]);
            }
        },

        /**
         * Invoke the registered global error handler, or fall back to console.error.
         * Page developers can register a handler via G3AxonLive.setErrorHandler(fn).
         * @param {Error} error - The error to report
         */
        onError: function (error) {
            if (window.G3AxonLiveOnError && typeof window.G3AxonLiveOnError === 'function') {
                window.G3AxonLiveOnError(error);
            } else {
                console.error('G3AxonLive error:', error);
            }
        }
    };

    /**
     * Public API: Register a custom error handler for all G3AxonLive errors.
     * Usage: G3AxonLive.setErrorHandler(function(err) { alert(err.message); });
     * @param {Function} handler - Function receiving an Error object
     */
    G3AxonLive.setErrorHandler = function (handler) {
        window.G3AxonLiveOnError = handler;
    };

    /**
     * Public API: Manually trigger a component event from JavaScript.
     * Useful for programmatic control outside of DOM event binding.
     * Usage: G3AxonLive.trigger('myButton', 'onclick', { step: '1' });
     * @param {string} componentId - Target component ID
     * @param {string} eventName   - Event name to fire (e.g. "onclick")
     * @param {Object} eventArgs   - Optional key/value event arguments
     */
    G3AxonLive.trigger = function (componentId, eventName, eventArgs) {
        this.sendEvent(componentId, eventName, eventArgs || {});
    };

})();