package server

import (
	"strings"
	"sync"
)

// normalizeKey converts a key to lowercase for case-insensitive storage
func normalizeKey(key string) string {
	return strings.ToLower(key)
}

// ApplicationObject represents the ASP Application object
// Provides application-wide state with Lock/Unlock for thread safety
type ApplicationObject struct {
	// Contents collection (standard variables)
	contents map[string]interface{}
	// StaticObjects collection (objects declared in global.asa with <OBJECT>)
	staticObjects map[string]interface{}
	// Mutex for Lock/Unlock methods
	mutex sync.RWMutex
	// Lock state tracking
	locked bool
}

// NewApplicationObject creates a new Application object
func NewApplicationObject() *ApplicationObject {
	return &ApplicationObject{
		contents:      make(map[string]interface{}),
		staticObjects: make(map[string]interface{}),
		locked:        false,
	}
}

// Lock prevents other clients from modifying Application variables
func (app *ApplicationObject) Lock() {
	app.mutex.Lock()
	app.locked = true
}

// Unlock allows other clients to modify Application variables
func (app *ApplicationObject) Unlock() {
	if app.locked {
		app.locked = false
		app.mutex.Unlock()
	}
}

// Get retrieves a value from Application.Contents (case-insensitive)
func (app *ApplicationObject) Get(key string) interface{} {
	app.mutex.RLock()
	defer app.mutex.RUnlock()
	
	// Normalize key to lowercase for case-insensitive lookup
	keyLower := normalizeKey(key)
	return app.contents[keyLower]
}

// Set stores a value in Application.Contents (case-insensitive)
func (app *ApplicationObject) Set(key string, value interface{}) {
	// Use RLock for read check, then upgrade if needed
	app.mutex.Lock()
	defer app.mutex.Unlock()
	
	// Normalize key to lowercase for case-insensitive storage
	keyLower := normalizeKey(key)
	app.contents[keyLower] = value
}

// Remove removes a key from Application.Contents
func (app *ApplicationObject) Remove(key string) {
	app.mutex.Lock()
	defer app.mutex.Unlock()
	
	keyLower := normalizeKey(key)
	delete(app.contents, keyLower)
}

// RemoveAll clears all variables from Application.Contents
func (app *ApplicationObject) RemoveAll() {
	app.mutex.Lock()
	defer app.mutex.Unlock()
	
	app.contents = make(map[string]interface{})
}

// GetContents returns the Contents collection for enumeration
// Returns a copy to avoid concurrent modification issues
func (app *ApplicationObject) GetContents() map[string]interface{} {
	app.mutex.RLock()
	defer app.mutex.RUnlock()
	
	// Return a copy of contents
	contentsCopy := make(map[string]interface{})
	for k, v := range app.contents {
		contentsCopy[k] = v
	}
	return contentsCopy
}

// GetStaticObjects returns the StaticObjects collection for enumeration
// For backwards compatibility, returns all keys (Contents + StaticObjects)
// Returns a copy to avoid concurrent modification issues
func (app *ApplicationObject) GetStaticObjects() map[string]interface{} {
	app.mutex.RLock()
	defer app.mutex.RUnlock()
	
	// Return combined map of both contents and static objects
	combined := make(map[string]interface{})
	for k, v := range app.contents {
		combined[k] = v
	}
	for k, v := range app.staticObjects {
		combined[k] = v
	}
	return combined
}

// AddStaticObject adds an object to the StaticObjects collection
// This is typically called when parsing global.asa <OBJECT> tags
func (app *ApplicationObject) AddStaticObject(key string, obj interface{}) {
	app.mutex.Lock()
	defer app.mutex.Unlock()
	
	keyLower := normalizeKey(key)
	app.staticObjects[keyLower] = obj
}

// GetAllKeys returns all keys from both Contents and StaticObjects
// Used for variable enumeration (For Each)
func (app *ApplicationObject) GetAllKeys() []string {
	app.mutex.RLock()
	defer app.mutex.RUnlock()
	
	keys := make([]string, 0, len(app.contents)+len(app.staticObjects))
	
	// Add contents keys
	for k := range app.contents {
		keys = append(keys, k)
	}
	
	// Add static objects keys
	for k := range app.staticObjects {
		keys = append(keys, k)
	}
	
	return keys
}

// Count returns the total number of items in Contents and StaticObjects
func (app *ApplicationObject) Count() int {
	app.mutex.RLock()
	defer app.mutex.RUnlock()
	
	return len(app.contents) + len(app.staticObjects)
}
