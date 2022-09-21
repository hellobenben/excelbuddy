package validator

import (
	"errors"
	"fmt"
	"regexp"
)

type RegExpValidator struct {
	Pattern string
}

func (r RegExpValidator) Validate(v string) error {
	ok, err := regexp.MatchString(r.Pattern, v)
	if err != nil {
		return err
	}
	if !ok {
		return errors.New(fmt.Sprintf("%s can't match the pattern %s", v, r.Pattern))
	}
	return nil
}

type Required struct {
}

func (r Required) Validate(v string) error {
	if len(v) == 0 {
		return errors.New("this value can't be empty")
	}
	return nil
}

type SnValidator struct {
}

func (r SnValidator) Validate(v string) error {
	if len(v) == 0 {
		return errors.New("this value can't be empty")
	}
	return nil
}
