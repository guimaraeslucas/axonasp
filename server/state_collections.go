package server

import (
	"fmt"
	"strings"
)

// SessionContentsCollection represents Session.Contents in classic ASP.
type SessionContentsCollection struct {
	session *SessionObject
}

func NewSessionContentsCollection(session *SessionObject) *SessionContentsCollection {
	return &SessionContentsCollection{session: session}
}

func (c *SessionContentsCollection) GetName() string {
	return "Session.Contents"
}

func (c *SessionContentsCollection) GetProperty(name string) interface{} {
	if c == nil || c.session == nil {
		return nil
	}

	switch strings.ToLower(name) {
	case "count":
		return c.session.Count()
	case "keys":
		return c.session.GetAllKeys()
	}

	return nil
}

func (c *SessionContentsCollection) SetProperty(name string, value interface{}) error {
	if c == nil || c.session == nil {
		return nil
	}
	return c.session.SetProperty(name, value)
}

func (c *SessionContentsCollection) CallMethod(name string, args ...interface{}) (interface{}, error) {
	if c == nil || c.session == nil {
		return nil, nil
	}

	switch strings.ToLower(name) {
	case "removeall":
		c.session.RemoveAll()
		return nil, nil
	case "remove":
		if len(args) == 0 {
			return nil, nil
		}
		key := fmt.Sprintf("%v", args[0])
		c.session.Remove(key)
		return nil, nil
	case "item":
		if len(args) == 0 {
			return nil, nil
		}
		return c.session.GetIndex(args[0]), nil
	case "count":
		return c.session.Count(), nil
	}

	return nil, nil
}

func (c *SessionContentsCollection) Enumerate() ([]interface{}, error) {
	if c == nil || c.session == nil {
		return []interface{}{}, nil
	}
	keys := c.session.GetAllKeys()
	items := make([]interface{}, 0, len(keys))
	for _, key := range keys {
		items = append(items, key)
	}
	return items, nil
}

// ApplicationContentsCollection represents Application.Contents in classic ASP.
type ApplicationContentsCollection struct {
	application *ApplicationObject
}

func NewApplicationContentsCollection(application *ApplicationObject) *ApplicationContentsCollection {
	return &ApplicationContentsCollection{application: application}
}

func (c *ApplicationContentsCollection) GetName() string {
	return "Application.Contents"
}

func (c *ApplicationContentsCollection) GetProperty(name string) interface{} {
	if c == nil || c.application == nil {
		return nil
	}

	switch strings.ToLower(name) {
	case "count":
		return len(c.application.GetContentKeys())
	case "keys":
		return c.application.GetContentKeys()
	}

	return nil
}

func (c *ApplicationContentsCollection) SetProperty(name string, value interface{}) error {
	if c == nil || c.application == nil {
		return nil
	}
	c.application.Set(name, value)
	return nil
}

func (c *ApplicationContentsCollection) CallMethod(name string, args ...interface{}) (interface{}, error) {
	if c == nil || c.application == nil {
		return nil, nil
	}

	switch strings.ToLower(name) {
	case "removeall":
		c.application.RemoveAll()
		return nil, nil
	case "remove":
		if len(args) == 0 {
			return nil, nil
		}
		c.application.Remove(fmt.Sprintf("%v", args[0]))
		return nil, nil
	case "item":
		if len(args) == 0 {
			return nil, nil
		}
		return c.application.Get(fmt.Sprintf("%v", args[0])), nil
	case "count":
		return len(c.application.GetContentKeys()), nil
	}

	return nil, nil
}

func (c *ApplicationContentsCollection) Enumerate() ([]interface{}, error) {
	if c == nil || c.application == nil {
		return []interface{}{}, nil
	}
	keys := c.application.GetContentKeys()
	items := make([]interface{}, 0, len(keys))
	for _, key := range keys {
		items = append(items, key)
	}
	return items, nil
}
