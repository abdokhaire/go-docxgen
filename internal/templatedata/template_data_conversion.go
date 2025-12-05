package templatedata

import (
	"fmt"
	"reflect"
)

// DataToMap converts any data type to a map[string]any for template processing.
// Supports structs, maps, pointers, slices of any type.
func DataToMap(data any) (map[string]any, error) {
	if data == nil {
		return nil, fmt.Errorf("data is nil")
	}

	if m, ok := data.(map[string]any); ok {
		return deepCopyMap(m), nil
	}

	return convertToMap(reflect.ValueOf(data))
}

// convertToMap recursively converts a reflect.Value to a map[string]any
func convertToMap(val reflect.Value) (map[string]any, error) {
	// Dereference pointers
	val = dereferenceValue(val)

	if !val.IsValid() {
		return nil, fmt.Errorf("invalid value (nil pointer)")
	}

	switch val.Kind() {
	case reflect.Struct:
		return convertStructToMap(val)
	case reflect.Map:
		return convertMapToMap(val)
	default:
		return nil, fmt.Errorf("expected a struct or map, got %s", val.Kind())
	}
}

// dereferenceValue dereferences pointers and interfaces to get the underlying value
func dereferenceValue(val reflect.Value) reflect.Value {
	for val.Kind() == reflect.Ptr || val.Kind() == reflect.Interface {
		if val.IsNil() {
			return reflect.Value{}
		}
		val = val.Elem()
	}
	return val
}

// convertStructToMap converts a struct to map[string]any
func convertStructToMap(val reflect.Value) (map[string]any, error) {
	result := make(map[string]any)

	for i := range val.NumField() {
		field := val.Type().Field(i)

		// Skip unexported fields
		if !field.IsExported() {
			continue
		}

		fieldValue := val.Field(i)
		converted, err := convertValue(fieldValue)
		if err != nil {
			return nil, fmt.Errorf("field %s: %w", field.Name, err)
		}
		result[field.Name] = converted
	}

	return result, nil
}

// convertMapToMap converts a map to map[string]any
func convertMapToMap(val reflect.Value) (map[string]any, error) {
	result := make(map[string]any)

	iter := val.MapRange()
	for iter.Next() {
		key := iter.Key()
		mapValue := iter.Value()

		// Convert key to string
		keyStr, err := keyToString(key)
		if err != nil {
			return nil, err
		}

		converted, err := convertValue(mapValue)
		if err != nil {
			return nil, fmt.Errorf("key %s: %w", keyStr, err)
		}
		result[keyStr] = converted
	}

	return result, nil
}

// keyToString converts a map key to a string
func keyToString(key reflect.Value) (string, error) {
	key = dereferenceValue(key)
	switch key.Kind() {
	case reflect.String:
		return key.String(), nil
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return fmt.Sprintf("%d", key.Int()), nil
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return fmt.Sprintf("%d", key.Uint()), nil
	default:
		return fmt.Sprintf("%v", key.Interface()), nil
	}
}

// convertValue converts any value to a template-compatible type
func convertValue(val reflect.Value) (any, error) {
	val = dereferenceValue(val)

	if !val.IsValid() {
		return nil, nil // nil pointer becomes nil
	}

	switch val.Kind() {
	case reflect.Struct:
		return convertStructToMap(val)

	case reflect.Map:
		return convertMapToMap(val)

	case reflect.Slice, reflect.Array:
		return convertSlice(val)

	case reflect.String:
		return val.String(), nil

	case reflect.Bool:
		return val.Bool(), nil

	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return val.Int(), nil

	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return val.Uint(), nil

	case reflect.Float32, reflect.Float64:
		return val.Float(), nil

	case reflect.Interface:
		if val.IsNil() {
			return nil, nil
		}
		return convertValue(val.Elem())

	default:
		// For other types (func, chan, etc.), return as-is
		if val.CanInterface() {
			return val.Interface(), nil
		}
		return nil, nil
	}
}

// convertSlice converts a slice/array to []any or []map[string]any
func convertSlice(val reflect.Value) (any, error) {
	length := val.Len()

	// Check if it's a slice of structs or maps (convert to []map[string]any)
	if length > 0 {
		firstElem := dereferenceValue(val.Index(0))
		if firstElem.IsValid() && (firstElem.Kind() == reflect.Struct || firstElem.Kind() == reflect.Map) {
			mapSlice := make([]map[string]any, length)
			for i := range length {
				elem := val.Index(i)
				converted, err := convertValue(elem)
				if err != nil {
					return nil, fmt.Errorf("index %d: %w", i, err)
				}
				if m, ok := converted.(map[string]any); ok {
					mapSlice[i] = m
				} else {
					// If conversion didn't return a map, wrap it
					mapSlice[i] = map[string]any{"Value": converted}
				}
			}
			return mapSlice, nil
		}
	}

	// For slices of primitives or empty slices, return []any
	result := make([]any, length)
	for i := range length {
		elem := val.Index(i)
		converted, err := convertValue(elem)
		if err != nil {
			return nil, fmt.Errorf("index %d: %w", i, err)
		}
		result[i] = converted
	}
	return result, nil
}

// deepCopyMap creates a deep copy of a map[string]any
func deepCopyMap(m map[string]any) map[string]any {
	result := make(map[string]any, len(m))
	for k, v := range m {
		switch val := v.(type) {
		case map[string]any:
			result[k] = deepCopyMap(val)
		case []map[string]any:
			newSlice := make([]map[string]any, len(val))
			for i, item := range val {
				newSlice[i] = deepCopyMap(item)
			}
			result[k] = newSlice
		case []any:
			newSlice := make([]any, len(val))
			copy(newSlice, val)
			result[k] = newSlice
		default:
			result[k] = v
		}
	}
	return result
}
