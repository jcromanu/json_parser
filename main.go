package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"reflect"
	"strings"

	xlst "github.com/ivahaev/go-xlsx-templater"
)

func main() {
	var jsonMap map[string]interface{}
	data := sampleData()
	err := json.Unmarshal(data, &jsonMap)
	if err != nil {
		fmt.Println("Final log")
	}
	ctx := castMap(jsonMap)

	// OS ejecution
	wd, err := os.Getwd()
	if err != nil {
		log.Fatalf(err.Error())
	}
	path := strings.Join([]string{wd, ""}, "")
	doc := xlst.New()
	err = doc.ReadTemplate(path + "/export_support_template.xlsx")
	if err != nil {
		fmt.Println("ERROR OPENING THE TEMPLATE: ", err)
		panic("error opening template")
	}
	err = doc.Render(ctx)
	if err != nil {
		fmt.Println("ERROR RENDERING THE TEMPLATE: ", err)
		panic("error rendering template")
	}
	err = doc.Save(path + "/report.xlsx")
	if err != nil {
		fmt.Println("ERROR SAVING THE TEMPLATE: ", err)
		panic("error saving template")
	}
}

func parseJson(context map[string]interface{}) map[string]interface{} {
	var vendorsMapArray []map[string]interface{}
	for key, val := range context {
		if key == "Vendors" {
			for _, vvalue := range val.([]interface{}) {
				vendorsMapArray = append(vendorsMapArray, vvalue.(map[string]interface{}))
			}
			context["Vendors"] = vendorsMapArray
		}
	}
	for _, nval := range context["Vendors"].([]map[string]interface{}) {
		for i, xval := range nval {
			var methotOfTendersMapArray []map[string]interface{}
			if i == "MethodOfTenders" {
				for _, mvalue := range xval.([]interface{}) {
					methotOfTendersMapArray = append(methotOfTendersMapArray, mvalue.(map[string]interface{}))
				}
				nval[i] = methotOfTendersMapArray
			}
		}
	}
	return context
}

func sampleData() []byte {
	return []byte(`{
		"GrandTotalAverageOrder": 4.273141025641025e+01,
		"GrandTotalFees": 2.261999995899201e+02,
		"GrandTotalNetSales": 6.332115000000000e+03,
		"GrandTotalOrders": 142,
		"GrandTotalRevenue": 6.666099999999999e+03,
		"GrandTotalTax": 3.692400000000000e+02,
		"GrandTotalTips": 5.569000000000000e+01,
		"Vendors": [
		  {
			"MethodOfTenders": [
			  {
				"AverageOrder": 5.664333333333334e+01,
				"Fees": 6.000000000000000e+00,
				"MethodOfTender": "Visa",
				"NetSales": 1.522138708493854e+02,
				"Orders": 3,
				"Tax": 1.172000000000000e+01,
				"Tips": 0.000000000000000e+00,
				"Total": 1.699300000000000e+02,
				"VendorName": "Jonas Vendor - Station",
				"VenueName": "Playa Vista QA Stadium"
			  },
			  {
				"AverageOrder": 5.894888888888889e+01,
				"Fees": 2.617999958992004e+01,
				"MethodOfTender": "Cash",
				"NetSales": 4.741761291506146e+02,
				"Orders": 7,
				"Tax": 3.018000000000000e+01,
				"Tips": 0.000000000000000e+00,
				"Total": 5.305400000000000e+02,
				"VendorName": "Jonas Vendor - Station",
				"VenueName": "Playa Vista QA Stadium"
			  }
			],
			"SubTotalAverageOrder": 3.502350000000000e+02,
			"SubTotalFees": 3.217999958992004e+01,
			"SubTotalNetSales": 6.263900000000001e+02,
			"SubTotalOrders": 10,
			"SubTotalRevenue": 7.004700000000000e+02,
			"SubTotalTax": 4.190000000000000e+01,
			"SubTotalTips": 0.000000000000000e+00
		  },
		  {
			"MethodOfTenders": [
			  {
				"AverageOrder": 1.130463636363636e+02,
				"Fees": 5.257000000000000e+01,
				"MethodOfTender": "Visa",
				"NetSales": 3.591735697663717e+03,
				"Orders": 32,
				"Tax": 3.115624138931910e+01,
				"Tips": 5.508000000000000e+01,
				"Total": 3.730530000000000e+03,
				"VendorName": "Appetize Demo Vendor P1",
				"VenueName": "Playa Vista QA Stadium"
			  },
			  {
				"AverageOrder": 2.189829787234043e+01,
				"Fees": 1.264500000000000e+02,
				"MethodOfTender": "Cash",
				"NetSales": 1.674799302336283e+03,
				"Orders": 90,
				"Tax": 2.821637586106809e+02,
				"Tips": 0.000000000000000e+00,
				"Total": 2.058440000000000e+03,
				"VendorName": "Appetize Demo Vendor P1",
				"VenueName": "Playa Vista QA Stadium"
			  },
			  {
				"AverageOrder": 1.100000000000000e+01,
				"Fees": 0.000000000000000e+00,
				"MethodOfTender": "MasterCard",
				"NetSales": 1.100000000000000e+02,
				"Orders": 2,
				"Tax": 0.000000000000000e+00,
				"Tips": 0.000000000000000e+00,
				"Total": 2.200000000000000e+01,
				"VendorName": "Appetize Demo Vendor P1",
				"VenueName": "Playa Vista QA Stadium"
			  },
			  {
				"AverageOrder": 0.000000000000000e+00,
				"Fees": 0.000000000000000e+00,
				"MethodOfTender": "NO PAYMENT",
				"NetSales": 1.980600000000000e+02,
				"Orders": 3,
				"Tax": 0.000000000000000e+00,
				"Tips": 0.000000000000000e+00,
				"Total": 0.000000000000000e+00,
				"VendorName": "Appetize Demo Vendor P1",
				"VenueName": "Playa Vista QA Stadium"
			  },
			  {
				"AverageOrder": 1.610000000000000e+01,
				"Fees": 0.000000000000000e+00,
				"MethodOfTender": "American Express",
				"NetSales": 6.000000000000000e+01,
				"Orders": 4,
				"Tax": 4.400000000000000e+00,
				"Tips": 0.000000000000000e+00,
				"Total": 6.440000000000001e+01,
				"VendorName": "Appetize Demo Vendor P1",
				"VenueName": "Playa Vista QA Stadium"
			  },
			  {
				"AverageOrder": -6.099999999999999e+00,
				"Fees": 0.000000000000000e+00,
				"MethodOfTender": "On House",
				"NetSales": -2.210000000000001e+00,
				"Orders": 1,
				"Tax": 1.600000000000000e+00,
				"Tips": 6.100000000000000e-01,
				"Total": -6.099999999999999e+00,
				"VendorName": "Appetize Demo Vendor P1",
				"VenueName": "Playa Vista QA Stadium"
			  }
			],
			"SubTotalAverageOrder": 9.782116666666666e+02,
			"SubTotalFees": 1.790200000000000e+02,
			"SubTotalNetSales": 5.632385000000000e+03,
			"SubTotalOrders": 132,
			"SubTotalRevenue": 5.869270000000000e+03,
			"SubTotalTax": 3.193200000000001e+02,
			"SubTotalTips": 5.569000000000000e+01
		  },
		  {
			"MethodOfTenders": [
			  {
				"AverageOrder": 1.338400000000000e+01,
				"Fees": 1.500000000000000e+01,
				"MethodOfTender": "Cash",
				"NetSales": 4.644000000000000e+01,
				"Orders": 5,
				"Tax": 5.480000000000000e+00,
				"Tips": 0.000000000000000e+00,
				"Total": 6.692000000000000e+01,
				"VendorName": "Lee Vendor",
				"VenueName": "Playa Vista QA Stadium"
			  }
			],
			"SubTotalAverageOrder": 6.692000000000000e+01,
			"SubTotalFees": 1.500000000000000e+01,
			"SubTotalNetSales": 4.644000000000000e+01,
			"SubTotalOrders": 5,
			"SubTotalRevenue": 6.692000000000000e+01,
			"SubTotalTax": 5.480000000000000e+00,
			"SubTotalTips": 0.000000000000000e+00
		  },
		  {
			"MethodOfTenders": [
			  {
				"AverageOrder": 1.472000000000000e+01,
				"Fees": 0.000000000000000e+00,
				"MethodOfTender": "Cash",
				"NetSales": 2.690000000000000e+01,
				"Orders": 2,
				"Tax": 2.540000000000000e+00,
				"Tips": 0.000000000000000e+00,
				"Total": 2.944000000000000e+01,
				"VendorName": "Norman Vendor",
				"VenueName": "Playa Vista QA Stadium"
			  }
			],
			"SubTotalAverageOrder": 2.944000000000000e+01,
			"SubTotalFees": 0.000000000000000e+00,
			"SubTotalNetSales": 2.690000000000000e+01,
			"SubTotalOrders": 2,
			"SubTotalRevenue": 2.944000000000000e+01,
			"SubTotalTax": 2.540000000000000e+00,
			"SubTotalTips": 0.000000000000000e+00
		  }
		],
		"VenueId": 929,
		"VenueName": "Playa Vista QA Stadium"
	  }`)
}

func walk(v reflect.Value) {
	// Indirect through pointers and interfaces
	for v.Kind() == reflect.Ptr || v.Kind() == reflect.Interface {
		v = v.Elem()
	}
	switch v.Kind() {
	case reflect.Array, reflect.Slice:
		for i := 0; i < v.Len(); i++ {
			walk(v.Index(i))
		}
	case reflect.Map:
		for _, k := range v.MapKeys() {
			walk(v.MapIndex(k))
		}
	default:
		// handle other types
	}
}

// "Casts" map values to the desired type recursively
func castMap(m map[string]any) map[string]any {
	for k := range m {
		switch reflect.ValueOf(m[k]).Kind() {
		case reflect.Map:
			mm, ok := m[k].(map[string]any)
			if !ok {
				panic(fmt.Errorf("Expected map[string]any, got %T", m[k]))
			}
			m[k] = castMap(mm)
		case reflect.Slice, reflect.Array:
			ma, ok := m[k].([]any)
			if !ok {
				panic(fmt.Errorf("Expected []any, got %T", m[k]))
			}
			m[k] = castArray(ma)
		default:
			// fmt.Printf("%s: %T, kind %v\n", k, m[k], reflect.ValueOf(m[k]).Kind())
			continue
		}
	}
	return m
}

// "Casts" slice elements to the desired types recursively
func castArray(a []any) []map[string]any {
	res := []map[string]any{}
	for i := range a {
		switch reflect.ValueOf(a[i]).Kind() {
		case reflect.Map:
			am, ok := a[i].(map[string]any)
			if !ok {
				panic(fmt.Errorf("Expected map[string]any, got %T", a[i]))
			}
			am = castMap(am)
			res = append(res, am)
		default:
			panic(fmt.Errorf("Expected map[string]any, got %T", a[i]))
		}
	}
	return res
}
