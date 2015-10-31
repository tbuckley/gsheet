package gsheet

import (
	"encoding/xml"
	"fmt"
	"io/ioutil"
	"net/http"
)

type Link struct {
	Rel  string `xml:"rel,attr"`
	Href string `xml:"href,attr"`
	Type string `xml:"type,attr"`
}
type WorksheetEntry struct {
	ID    string `xml:"id"`
	Title string `xml:"title"`
	Links []Link `xml:"link"`
	Rows  int    `xml:"http://schemas.google.com/spreadsheets/2006 rowCount"`
	Cols  int    `xml:"http://schemas.google.com/spreadsheets/2006 colCount"`
}
type Spreadsheet struct {
	Title      string           `xml:"title"`
	Worksheets []WorksheetEntry `xml:"entry"`
}

type Cell struct {
	Row          int     `xml:"row,attr"`
	Col          int     `xml:"col,attr"`
	InputValue   string  `xml:"inputValue,attr"`
	NumericValue float64 `xml:"numericValue,attr"`
}
type Worksheet struct {
	Rows  int     `xml:"http://schemas.google.com/spreadsheets/2006 rowCount"`
	Cols  int     `xml:"http://schemas.google.com/spreadsheets/2006 colCount"`
	Cells []*Cell `xml:"entry>cell"`

	cellMap map[int]map[int]*Cell
}

type Service struct {
	client *http.Client
}

type SpreadsheetService struct {
	parent        *Service
	spreadsheetID string
}

type WorksheetService struct {
	parent      *SpreadsheetService
	worksheetID string
}

// Service

func New(client *http.Client) *Service {
	return &Service{
		client: client,
	}
}

func (svc *Service) Spreadsheet(ID string) *SpreadsheetService {
	return &SpreadsheetService{
		parent:        svc,
		spreadsheetID: ID,
	}
}

func (svc *SpreadsheetService) Get() (*Spreadsheet, error) {
	urlFormat := "https://spreadsheets.google.com/feeds/worksheets/%s/private/full"
	url := fmt.Sprintf(urlFormat, svc.spreadsheetID)
	resp, err := svc.parent.client.Get(url)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	data, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return nil, err
	}

	spreadsheet := new(Spreadsheet)
	err = xml.Unmarshal(data, &spreadsheet)
	if err != nil {
		return nil, err
	}
	return spreadsheet, nil
}

func (svc *SpreadsheetService) Worksheet(ID string) *WorksheetService {
	offset := len("https://spreadsheets.google.com/feeds/worksheets//private/full/") + len(svc.spreadsheetID)
	return &WorksheetService{
		parent:      svc,
		worksheetID: ID[offset:],
	}
}

func (svc *WorksheetService) Get() (*Worksheet, error) {
	urlFormat := "https://spreadsheets.google.com/feeds/cells/%v/%v/private/full"
	url := fmt.Sprintf(urlFormat, svc.parent.spreadsheetID, svc.worksheetID)
	resp, err := svc.parent.parent.client.Get(url)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	data, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return nil, err
	}

	worksheet := new(Worksheet)
	err = xml.Unmarshal(data, &worksheet)
	if err != nil {
		return nil, err
	}

	worksheet.process()

	return worksheet, nil
}

// Utilities

func (s *Spreadsheet) WorksheetIDByTitle(title string) (string, bool) {
	for _, worksheet := range s.Worksheets {
		if worksheet.Title == title {
			return worksheet.ID, true
		}
	}
	return "", false
}

func (w *Worksheet) process() {
	w.cellMap = make(map[int]map[int]*Cell)
	for _, cell := range w.Cells {
		if _, ok := w.cellMap[cell.Col]; !ok {
			w.cellMap[cell.Col] = make(map[int]*Cell)
		}
		w.cellMap[cell.Col][cell.Row] = cell
	}
}

func (w *Worksheet) Get(col, row int) (*Cell, bool) {
	rows, ok := w.cellMap[col]
	if !ok {
		return nil, false
	}

	cell, ok := rows[row]
	return cell, ok
}

func (w *Worksheet) GetColByTitle(title string) (int, bool) {
	for col, rows := range w.cellMap {
		if _, ok := rows[1]; ok {
			if rows[1].InputValue == title {
				return col, true
			}
		}
	}
	return 0, false
}
