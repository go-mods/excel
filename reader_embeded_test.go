package excel

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

type Named struct {
	ID   int    `excel:"Id" excel-in:"Id" excel-out:"id"`
	Name string `excel:"Name,default:error" excel-in:"default:anonymous" excel-out:"default:not_used"`
}

type NamedUser struct {
	Named
	Ignored     string    `excel:"-"`
	EncodedName Encoded   `excel:"Encoded_Name,encoding:json"`
	Created     time.Time `excel:"created,format:d/m/Y"`
	AnArray     []int     `excel:"array,split:;" excel-out:"split:|"`
}

func TestReadNamedUser(t *testing.T) {
	// Créer un fichier Excel pour le test
	inFile := excelize.NewFile()

	// Définir les en-têtes de colonnes
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "Id")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "Name")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "C1", "Encoded_Name")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "D1", "created")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "E1", "array")

	// Définir les valeurs
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "Test User")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "C2", "{\"name\":\"encoded name\"}")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "D2", "01/01/2023")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "E2", "1;2;3")

	defer func() { _ = inFile.Close() }()

	// Créer un conteneur pour les utilisateurs
	var namedUsers []NamedUser

	// Configurer le lecteur Excel
	inExcel, _ := NewReader(inFile)
	inExcel.SetSheet(inExcel.GetActiveSheet())
	inExcel.SetAxis("A1")

	// Désérialiser les données
	err := inExcel.Unmarshal(&namedUsers)
	if err != nil {
		t.Error(err)
		return
	}

	// Vérifier les résultats
	assert.Equal(t, 1, len(namedUsers), "Il devrait y avoir un utilisateur")
	assert.Equal(t, 1, namedUsers[0].ID, "L'ID devrait être 1")
	assert.Equal(t, "Test User", namedUsers[0].Name, "Le nom devrait être 'Test User'")
	assert.Equal(t, "encoded name", namedUsers[0].EncodedName.Name, "Le nom encodé devrait être 'encoded name'")
	assert.Equal(t, 3, len(namedUsers[0].AnArray), "Le tableau devrait contenir 3 éléments")
	assert.Equal(t, 1, namedUsers[0].AnArray[0], "Le premier élément du tableau devrait être 1")
	assert.Equal(t, 2, namedUsers[0].AnArray[1], "Le deuxième élément du tableau devrait être 2")
	assert.Equal(t, 3, namedUsers[0].AnArray[2], "Le troisième élément du tableau devrait être 3")
}

// func TestWriteNamedUser(t *testing.T) {
// 	// Créer un fichier Excel pour le test
// 	outFile := excelize.NewFile()
// 	defer func() { _ = outFile.Close() }()
//
// 	// Créer des données de test
// 	createdDate, _ := time.Parse("02/01/2006", "01/01/2023")
// 	namedUsers := []NamedUser{
// 		{
// 			Named: Named{
// 				ID:   1,
// 				Name: "Test User",
// 			},
// 			Ignored:     "This should be ignored",
// 			EncodedName: Encoded{Name: "encoded name"},
// 			Created:     createdDate,
// 			AnArray:     []int{1, 2, 3},
// 		},
// 	}
//
// 	// Configurer l'écrivain Excel
// 	outExcel, _ := NewWriter(outFile)
// 	outExcel.SetSheet(outExcel.GetActiveSheet())
// 	outExcel.SetAxis("A1")
//
// 	// Sérialiser les données
// 	err := outExcel.Marshal(&namedUsers)
// 	if err != nil {
// 		t.Error(err)
// 		return
// 	}
//
// 	// Vérifier les en-têtes
// 	headerA1, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "A1")
// 	headerB1, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "B1")
// 	headerC1, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "C1")
// 	headerD1, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "D1")
// 	headerE1, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "E1")
//
// 	assert.Equal(t, "id", headerA1, "L'en-tête A1 devrait être 'id'")
// 	assert.Equal(t, "Name", headerB1, "L'en-tête B1 devrait être 'Name'")
// 	assert.Equal(t, "Encoded_Name", headerC1, "L'en-tête C1 devrait être 'Encoded_Name'")
// 	assert.Equal(t, "created", headerD1, "L'en-tête D1 devrait être 'created'")
// 	assert.Equal(t, "array", headerE1, "L'en-tête E1 devrait être 'array'")
//
// 	// Vérifier les valeurs
// 	valueA2, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "A2")
// 	valueB2, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "B2")
// 	valueC2, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "C2")
// 	valueD2, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "D2")
// 	valueE2, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "E2")
//
// 	assert.Equal(t, "1", valueA2, "La valeur A2 devrait être '1'")
// 	assert.Equal(t, "Test User", valueB2, "La valeur B2 devrait être 'Test User'")
// 	assert.Equal(t, "{\"name\":\"encoded name\"}", valueC2, "La valeur C2 devrait être '{\"name\":\"encoded name\"}'")
// 	assert.Equal(t, "01/01/2023", valueD2, "La valeur D2 devrait être '01/01/2023'")
// 	assert.Equal(t, "1|2|3", valueE2, "La valeur E2 devrait être '1|2|3'")
// }
//
// func TestReadWriteNamedUser(t *testing.T) {
// 	// Créer un fichier Excel pour le test d'écriture
// 	outFile := excelize.NewFile()
//
// 	// Créer des données de test
// 	createdDate, _ := time.Parse("02/01/2006", "01/01/2023")
// 	originalUsers := []NamedUser{
// 		{
// 			Named: Named{
// 				ID:   1,
// 				Name: "Test User",
// 			},
// 			Ignored:     "This should be ignored",
// 			EncodedName: Encoded{Name: "encoded name"},
// 			Created:     createdDate,
// 			AnArray:     []int{1, 2, 3},
// 		},
// 	}
//
// 	// Configurer l'écrivain Excel
// 	outExcel, _ := NewWriter(outFile)
// 	outExcel.SetSheet(outExcel.GetActiveSheet())
// 	outExcel.SetAxis("A1")
//
// 	// Sérialiser les données
// 	err := outExcel.Marshal(&originalUsers)
// 	if err != nil {
// 		t.Error(err)
// 		return
// 	}
//
// 	// Maintenant lire les données du même fichier
// 	var readUsers []NamedUser
//
// 	// Configurer le lecteur Excel
// 	inExcel, _ := NewReader(outFile)
// 	inExcel.SetSheet(inExcel.GetActiveSheet())
// 	inExcel.SetAxis("A1")
//
// 	// Désérialiser les données
// 	err = inExcel.Unmarshal(&readUsers)
// 	if err != nil {
// 		t.Error(err)
// 		return
// 	}
//
// 	defer func() { _ = outFile.Close() }()
//
// 	// Vérifier que les données lues correspondent aux données écrites
// 	assert.Equal(t, 1, len(readUsers), "Il devrait y avoir un utilisateur")
// 	assert.Equal(t, originalUsers[0].ID, readUsers[0].ID, "Les IDs devraient correspondre")
// 	assert.Equal(t, originalUsers[0].Name, readUsers[0].Name, "Les noms devraient correspondre")
// 	assert.Equal(t, originalUsers[0].EncodedName.Name, readUsers[0].EncodedName.Name, "Les noms encodés devraient correspondre")
// 	assert.Equal(t, len(originalUsers[0].AnArray), len(readUsers[0].AnArray), "Les tableaux devraient avoir la même taille")
// 	for i := 0; i < len(originalUsers[0].AnArray); i++ {
// 		assert.Equal(t, originalUsers[0].AnArray[i], readUsers[0].AnArray[i], "Les éléments du tableau devraient correspondre")
// 	}
// }
