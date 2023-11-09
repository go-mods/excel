package excel

import (
	"fmt"
	"github.com/stretchr/testify/assert"
	"testing"
)

func mustRange(ref string) *Range {
	r, err := ToRange(ref)
	if err != nil {
		panic(err)
	}
	return r
}

func TestToRange(t *testing.T) {
	type args struct{ ref string }
	type test struct {
		name    string
		args    args
		want    *Range
		wantErr assert.ErrorAssertionFunc
	}

	tests := []test{
		{
			name:    "ToRange(A1:B2)",
			args:    args{ref: "A1:B2"},
			want:    &Range{StartColumn: 1, StartRow: 1, StartName: "A1", EndColumn: 2, EndRow: 2, EndName: "B2"},
			wantErr: assert.NoError,
		},
		{
			name:    "ToRange($A$1:$B$2)",
			args:    args{ref: "A1:B2"},
			want:    &Range{StartColumn: 1, StartRow: 1, StartName: "A1", EndColumn: 2, EndRow: 2, EndName: "B2"},
			wantErr: assert.NoError,
		},
		{
			name:    "ToRange(A1)",
			args:    args{ref: "A1"},
			want:    nil,
			wantErr: assert.Error,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := ToRange(tt.args.ref)
			if !tt.wantErr(t, err, fmt.Sprintf("ToRange(%v)", tt.args.ref)) {
				return
			}
			assert.Equalf(t, tt.want, got, "ToRange(%v)", tt.args.ref)
		})
	}
}

func TestRange_ToRef(t *testing.T) {
	type args struct{ ref *Range }
	type test struct {
		name    string
		args    args
		want    string
		wantErr assert.ErrorAssertionFunc
	}

	tests := []test{
		{
			name:    "ToRef(A1:B2)",
			args:    args{mustRange("A1:B2")},
			want:    "A1:B2",
			wantErr: assert.NoError,
		},
		{
			name:    "ToRef($A$1:$B$2)",
			args:    args{mustRange("$A$1:$B$2")},
			want:    "A1:B2",
			wantErr: assert.NoError,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := tt.args.ref.ToRef()
			assert.Equalf(t, tt.want, got, "ToRef(%v)", tt.args.ref)
		})
	}

}

func TestRows(t *testing.T) {
	A1B2 := mustRange("A1:B2")

	assert.Equalf(t, A1B2.Rows(), 2, "Rows(%v)", A1B2)
	assert.Equalf(t, A1B2.ToRef(), "A1:B2", "ToRef(%v)", A1B2)

	assert.NoError(t, A1B2.AddRows(1))
	assert.Equalf(t, A1B2.Rows(), 3, "Rows(%v)", A1B2)
	assert.Equalf(t, A1B2.ToRef(), "A1:B3", "ToRef(%v)", A1B2)

	assert.NoError(t, A1B2.AddRows(-1))
	assert.Equalf(t, A1B2.Rows(), 2, "Rows(%v)", A1B2)
	assert.Equalf(t, A1B2.ToRef(), "A1:B2", "ToRef(%v)", A1B2)

	assert.NoError(t, A1B2.RemoveRows(1))
	assert.Equalf(t, A1B2.Rows(), 1, "Rows(%v)", A1B2)
	assert.Equalf(t, A1B2.ToRef(), "A1:B1", "ToRef(%v)", A1B2)

	assert.NoError(t, A1B2.SetRows(2))
	assert.Equalf(t, A1B2.Rows(), 2, "Rows(%v)", A1B2)
	assert.Equalf(t, A1B2.ToRef(), "A1:B2", "ToRef(%v)", A1B2)
}

func TestColumns(t *testing.T) {
	A1B2 := mustRange("A1:B2")

	assert.Equalf(t, A1B2.Columns(), 2, "Columns(%v)", A1B2)
	assert.Equalf(t, A1B2.ToRef(), "A1:B2", "ToRef(%v)", A1B2)

	assert.NoError(t, A1B2.AddColumns(1))
	assert.Equalf(t, A1B2.Columns(), 3, "Columns(%v)", A1B2)
	assert.Equalf(t, A1B2.ToRef(), "A1:C2", "ToRef(%v)", A1B2)

	assert.NoError(t, A1B2.AddColumns(-1))
	assert.Equalf(t, A1B2.Columns(), 2, "Columns(%v)", A1B2)
	assert.Equalf(t, A1B2.ToRef(), "A1:B2", "ToRef(%v)", A1B2)

	assert.NoError(t, A1B2.RemoveColumns(1))
	assert.Equalf(t, A1B2.Columns(), 1, "Columns(%v)", A1B2)
	assert.Equalf(t, A1B2.ToRef(), "A1:A2", "ToRef(%v)", A1B2)

	assert.NoError(t, A1B2.SetColumns(2))
	assert.Equalf(t, A1B2.Columns(), 2, "Columns(%v)", A1B2)
	assert.Equalf(t, A1B2.ToRef(), "A1:B2", "ToRef(%v)", A1B2)
}
