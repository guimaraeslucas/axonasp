/**
 * G3AxonLive Client Engine v2.0
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * G3Pix AxonASP JavaScript engine for reactive ASP components.
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
        // Optional legacy session ID kept for compatibility with existing page templates
        sessionId: null,

        // Tracks whether an async fetch operation is currently in flight (debounce guard)
        isProcessing: false,

        // Delegated handlers are attached to document only once per page.
        handlersAttached: false,

        // Endpoint URL for all G3AxonLive fetch requests
        endpoint: '/g3al',

        // Retry configuration for exponential backoff on transient network errors
        maxRetries: 3,
        retryBaseDelayMs: 1000,

        /**
         * Initialize the G3AxonLive engine on page load.
         * The sessionId argument is optional and retained for backward compatibility.
         * @param {string} sessionId - The user's ASP Session ID
         * @returns {boolean} - True if initialization succeeded
         */
        init: function (sessionId) {
            this.sessionId = sessionId || null;
            this.attachComponentEventHandlers();
            return true;
        },

        /**
         * Attach delegated click, change, and submit event handlers at the document level.
         * Reactive components are identified by the data-g3al-id attribute.
         * Event type is specified by the data-g3al-event attribute (click, change, submit).
         */
        attachComponentEventHandlers: function () {
            if (this.handlersAttached) {
                return;
            }

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

            this.handlersAttached = true;
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
         * Also automatically includes the value of form elements.
         * @param {HTMLElement} component - The reactive component element
         * @returns {Object} - Map of argument names to string values
         */
        extractEventArgs: function (component) {
            var args = {};
            if (!component.attributes) return args;

            // 1. Existing data-g3al-arg-* extraction
            for (var i = 0; i < component.attributes.length; i++) {
                var attr = component.attributes[i];
                if (attr.name.indexOf('data-g3al-arg-') === 0) {
                    var argName = attr.name.substring('data-g3al-arg-'.length);
                    args[argName] = attr.value;
                }
            }

            // 2. Automatic value extraction for form elements
            var tagName = component.tagName.toLowerCase();
            var type = (component.type || '').toLowerCase();

            if (tagName === 'input' || tagName === 'select' || tagName === 'textarea') {
                if (type === 'radio') {
                    // Handle Radio Group: if a name is present, find the checked one in the same group
                    var name = component.getAttribute('name');
                    if (name) {
                        var radios = document.getElementsByName(name);
                        for (var j = 0; j < radios.length; j++) {
                            if (radios[j].checked) {
                                args['value'] = radios[j].value;
                                break;
                            }
                        }
                    } else {
                        args['value'] = component.checked ? component.value : '';
                    }
                } else if (type === 'checkbox') {
                    args['value'] = component.checked ? component.value : '';
                    args['checked'] = component.checked ? 'true' : 'false';
                } else if (tagName === 'select' && component.multiple) {
                    var selectedValues = [];
                    for (var s = 0; s < component.options.length; s++) {
                        if (component.options[s].selected) selectedValues.push(component.options[s].value);
                    }
                    args['value'] = selectedValues.join(',');
                } else {
                    args['value'] = component.value;
                }
            } else if (component.getAttribute('data-g3al-type') === 'checkboxlist') {
                // Special handling for a container representing a checkbox list
                var selected = [];
                var cbks = component.querySelectorAll('input[type="checkbox"]');
                for (var k = 0; k < cbks.length; k++) {
                    if (cbks[k].checked) selected.push(cbks[k].value);
                }
                args['value'] = selected.join(',');
            } else if (component.getAttribute('data-g3al-type') === 'radiobuttonlist') {
                // Special handling for a container representing a radio button list
                var checkedRadio = component.querySelector('input[type="radio"]:checked');
                args['value'] = checkedRadio ? checkedRadio.value : '';
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
            var self = this;
            this.isProcessing = true;

            var payload = {
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

                        case 'set_property':
                            // Set a property on the identified component element (e.g. value, disabled).
                            if (action.componentId && action.name) {
                                var el = document.getElementById(action.componentId);
                                if (el) {
                                    // Security constraint: block dangerous properties to prevent XSS.
                                    var blocked = ['innerHTML', 'outerHTML', 'onclick', 'onchange', 'onsubmit', 'onmouseover'];
                                    if (blocked.indexOf(action.name.toLowerCase()) !== -1 || action.name.indexOf('on') === 0) {
                                        console.warn('G3AxonLive: set_property blocked dangerous property:', action.name);
                                        break;
                                    }
                                    // Handle boolean conversion if necessary
                                    var val = action.value;
                                    if (val === 'true') val = true;
                                    if (val === 'false') val = false;
                                    el[action.name] = val;
                                } else {
                                    console.warn('G3AxonLive: set_property — element not found:', action.componentId);
                                }
                            }
                            break;

                        case 'set_style':
                            // Update a specific CSS style property.
                            if (action.componentId && action.name) {
                                var el = document.getElementById(action.componentId);
                                if (el) {
                                    el.style[action.name] = action.value || '';
                                } else {
                                    console.warn('G3AxonLive: set_style — element not found:', action.componentId);
                                }
                            }
                            break;

                        case 'add_class':
                            // Add a CSS class via classList.
                            if (action.componentId && action.value) {
                                var el = document.getElementById(action.componentId);
                                if (el) {
                                    el.classList.add(action.value);
                                } else {
                                    console.warn('G3AxonLive: add_class — element not found:', action.componentId);
                                }
                            }
                            break;

                        case 'remove_class':
                            // Remove a CSS class via classList.
                            if (action.componentId && action.value) {
                                var el = document.getElementById(action.componentId);
                                if (el) {
                                    el.classList.remove(action.value);
                                } else {
                                    console.warn('G3AxonLive: remove_class — element not found:', action.componentId);
                                }
                            }
                            break;

                        case 'remove_attribute':
                            // Remove an HTML attribute.
                            if (action.componentId && action.name) {
                                var el = document.getElementById(action.componentId);
                                if (el) {
                                    el.removeAttribute(action.name);
                                } else {
                                    console.warn('G3AxonLive: remove_attribute — element not found:', action.componentId);
                                }
                            }
                            break;

                        case 'add_title':
                            // Update the element's title (tooltip).
                            if (action.componentId) {
                                var el = document.getElementById(action.componentId);
                                if (el) {
                                    el.title = action.value || '';
                                }
                            }
                            break;

                        case 'remove_title':
                            // Clear the element's title.
                            if (action.componentId) {
                                var el = document.getElementById(action.componentId);
                                if (el) {
                                    el.title = '';
                                }
                            }
                            break;

                        case 'set_value':
                            // Direct value assignment (helper for form fields).
                            if (action.componentId) {
                                var el = document.getElementById(action.componentId);
                                if (el) {
                                    el.value = action.value || '';
                                }
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

    /**
     * Shows a modal component by ID.
     * @param {string} modalId - The DOM ID of the modal container.
     */
    G3AxonLive.showModal = function (modalId) {
        var el = document.getElementById(modalId);
        if (el) el.style.display = 'block';
    };

    /**
     * Closes a modal component by ID.
     * @param {string} modalId - The DOM ID of the modal container.
     */
    G3AxonLive.closeModal = function (modalId) {
        var el = document.getElementById(modalId);
        if (el) el.style.display = 'none';
    };

    /**
     * Toggles a modal component's visibility.
     * @param {string} modalId - The DOM ID of the modal container.
     */
    G3AxonLive.toggleModal = function (modalId) {
        var el = document.getElementById(modalId);
        if (el) {
            el.style.display = (el.style.display === 'none' || el.style.display === '') ? 'block' : 'none';
        }
    };

    /**
     * Handles a file upload using a specialized fetch request.
     * Sends multipart/form-data to the server and processes the reactive response.
     * @param {string} componentId - The target uploader component ID
     * @param {string} fileInputId - The ID of the input[type=file]
     * @param {string} eventName   - The event name to fire (e.g. "onupload")
     */
    G3AxonLive.uploadFile = function (componentId, fileInputId, eventName) {
        var input = document.getElementById(fileInputId);
        if (!input || !input.files || input.files.length === 0) return;

        var formData = new FormData();
        formData.append('file', input.files[0]);
        formData.append('g3al_id', componentId);
        formData.append('g3al_event', eventName);

        this.isProcessing = true;
        var self = this;

        fetch(window.location.href, {
            method: 'POST',
            body: formData,
            headers: {
                'X-Requested-With': 'G3AxonLive',
                'X-G3AL-Upload': 'true'
            }
        })
            .then(function (res) { return res.json(); })
            .then(function (data) {
                self.isProcessing = false;
                self.processResponse(data);
            })
            .catch(function (err) {
                self.isProcessing = false;
                console.error('G3AxonLive: Upload failed:', err);
            });
    };

})();