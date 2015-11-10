package gsheet

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"text/template"
)

var FEED_TEMPLATE = `<feed xmlns="http://www.w3.org/2005/Atom"
    xmlns:batch="http://schemas.google.com/gdata/batch"
    xmlns:gs="http://schemas.google.com/spreadsheets/2006">
  {{ $baseURL := .BaseURL}}
  <id>{{$baseURL}}</id>
  {{ range $entry := .Entries }}
  <entry>
    <batch:id>R{{$entry.Row}}C{{$entry.Col}}</batch:id>
    <batch:operation type="update"/>
    <id>{{$baseURL}}/R{{$entry.Row}}C{{$entry.Col}}</id>
    <link rel="edit" type="application/atom+xml" href="{{$entry.EditLink}}"/>
    <gs:cell row="{{$entry.Row}}" col="{{$entry.Col}}" inputValue="{{$entry.InputValue}}"/>
   </entry>
   {{ end }}
</feed>`

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
	Links        []Link  `xml:"link"`
}
type Worksheet struct {
	ID    string  `xml:"id"`
	Rows  int     `xml:"http://schemas.google.com/spreadsheets/2006 rowCount"`
	Cols  int     `xml:"http://schemas.google.com/spreadsheets/2006 colCount"`
	Cells []*Cell `xml:"entry>cell"`
	Links []Link  `xml:"link"`

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

type CellService struct {
	parent *WorksheetService
	row    int
	col    int
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

func (svc *WorksheetService) Cell(row, col int) *CellService {
	return &CellService{
		parent: svc,
		row:    row,
		col:    col,
	}
}

type BatchChangeCell struct {
	BaseURL string
	Entries []ChangeCell
}
type ChangeCell struct {
	EditLink   string
	Col        int
	Row        int
	InputValue string
}

type CellNotFoundError struct {
	Row int
	Col int
}

func (c *CellNotFoundError) Error() string {
	return fmt.Sprintf("cell (%v, %v) not found", c.Row, c.Col)
}

func (svc *CellService) Get() (*Cell, error) {
	urlFormat := "https://spreadsheets.google.com/feeds/cells/%v/%v/private/full/R%vC%v"
	url := fmt.Sprintf(urlFormat, svc.parent.parent.spreadsheetID, svc.parent.worksheetID, svc.row, svc.col)
	resp, err := svc.parent.parent.parent.client.Get(url)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	data, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return nil, err
	}

	cell := new(Cell)
	err = xml.Unmarshal(data, &cell)
	if err != nil {
		return nil, err
	}

	return cell, nil
}

func (svc *CellService) Set(inputValue string) error {
	cell, _ := svc.Get()
	editLink := getLinkByRel("edit", cell.Links)
	fmt.Printf("edit link: %v\n", editLink)

	urlFmt := "https://spreadsheets.google.com/feeds/cells/%v/%v/private/full/batch"
	url := fmt.Sprintf(urlFmt, svc.parent.parent.spreadsheetID, svc.parent.worksheetID)

	bodyTpl, err := template.New("change_cell").Parse(FEED_TEMPLATE)
	if err != nil {
		log.Fatalf("Error processing template: %v", err)
	}

	buf := new(bytes.Buffer)
	baseURLFmt := "https://spreadsheets.google.com/feeds/cells/%v/%v/private/full"
	baseURL := fmt.Sprintf(baseURLFmt, svc.parent.parent.spreadsheetID, svc.parent.worksheetID)
	err = bodyTpl.Execute(buf, &BatchChangeCell{
		BaseURL: baseURL,
		Entries: []ChangeCell{
			ChangeCell{EditLink: editLink, Col: svc.col, Row: svc.row, InputValue: inputValue},
		},
	})
	if err != nil {
		log.Fatalf("Error executing template: %v", err)
	}

	xmlData := buf.Bytes()
	readBuf := bytes.NewBuffer(xmlData)

	req, err := http.NewRequest("POST", url, readBuf)
	if err != nil {
		log.Fatalf("Cannot create request: %v", err)
	}

	req.Header.Add("Content-Type", "text/xml")

	resp, err := svc.parent.parent.parent.client.Do(req)
	if err != nil {
		return err
	}
	_, err = ioutil.ReadAll(resp.Body)
	if err != nil {
		log.Fatalf("Cannot read response: %v", err)
	}
	return nil
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

func getLinkByRel(rel string, links []Link) string {
	for _, link := range links {
		if link.Rel == rel {
			return link.Href
		}
	}
	return ""
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

func (w *Worksheet) GetID() string {
	var key, worksheetID string
	fmt.Sscanf(w.ID, "https://spreadsheets.google.com/feeds/cells/%s/%s/private/full", &key, &worksheetID)
	return worksheetID
}
