package docx

import (
	"encoding/xml"
	"strings"
	"testing"
)

const xml2merge = `<w:p w14:paraId="343EA723" w14:textId="17A5316C" w:rsidR="00B7252F" w:rsidRPr="00334290" w:rsidRDefault="00B7252F" w:rsidP="00334290">
<w:pPr>
	<w:spacing w:after="120" w:line="240" w:lineRule="atLeast"/>
	<w:jc w:val="center"/>
	<w:rPr>
		<w:b/>
		<w:sz w:val="28"/>
		<w:u w:val="single"/>
	</w:rPr>
</w:pPr>
<w:r w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>某某某大学</w:t>
</w:r>
<w:r w:rsidR="00DC7F59" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>20</w:t>
</w:r>
<w:r w:rsidR="00F276CD" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>1</w:t>
</w:r>
<w:r w:rsidR="00AC3815">
	<w:rPr>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>6</w:t>
</w:r>
<w:r w:rsidR="00DC7F59" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>-20</w:t>
</w:r>
<w:r w:rsidR="00F276CD" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>1</w:t>
</w:r>
<w:r w:rsidR="00AC3815">
	<w:rPr>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>7</w:t>
</w:r>
<w:proofErr w:type="gramStart"/>
<w:r w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>学年第</w:t>
</w:r>
<w:proofErr w:type="gramEnd"/>
<w:r w:rsidR="007A75E1" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
		<w:u w:val="single"/>
	</w:rPr>
	<w:t xml:space="preserve"></w:t>
</w:r>
<w:r w:rsidR="00BA388C" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
		<w:u w:val="single"/>
	</w:rPr>
	<w:t>1</w:t>
</w:r>
<w:r w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
		<w:u w:val="single"/>
	</w:rPr>
	<w:t xml:space="preserve"></w:t>
</w:r>
<w:r w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>学期期</w:t>
</w:r>
<w:r w:rsidR="007A75E1" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
		<w:u w:val="single"/>
	</w:rPr>
	<w:t xml:space="preserve"></w:t>
</w:r>
<w:r w:rsidR="007A75E1" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
		<w:u w:val="single"/>
	</w:rPr>
	<w:t>末</w:t>
</w:r>
<w:r w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
		<w:u w:val="single"/>
	</w:rPr>
	<w:t xml:space="preserve"></w:t>
</w:r>
<w:r w:rsidR="006B05F0" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>考试</w:t>
</w:r>
<w:r w:rsidR="00DC7F59" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
		<w:u w:val="single"/>
	</w:rPr>
	<w:t xml:space="preserve"></w:t>
</w:r>
<w:r w:rsidR="00B75B37" w:rsidRPr="00027D88">
	<w:rPr>
		<w:b/>
		<w:sz w:val="28"/>
		<w:u w:val="single"/>
	</w:rPr>
	<w:t>A</w:t>
</w:r>
<w:r w:rsidR="00DC7F59" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
		<w:u w:val="single"/>
	</w:rPr>
	<w:t xml:space="preserve"></w:t>
</w:r>
<w:r w:rsidR="00DC7F59" w:rsidRPr="00027D88">
	<w:rPr>
		<w:rFonts w:hint="eastAsia"/>
		<w:b/>
		<w:sz w:val="28"/>
	</w:rPr>
	<w:t>卷</w:t>
</w:r>
</w:p>`

const (
	allmergedtext      = `某某某大学2016-2017学年第1学期期末考试A卷`
	propmergedtext     = `某某某大学201|6|-201|7|学年第|1|学期期|末|考试||A||卷|`
	namedpropmergdtext = `某某某大学2016-2017学年第|1|学期期|末|考试|A|卷|`
)

func TestMergeText(t *testing.T) {
	p := Paragraph{}
	err := xml.Unmarshal(StringToBytes(xml2merge), &p)
	if err != nil {
		t.Fatal(err)
	}
	np := p.MergeText(MergeAllRuns)
	// Count only Run children (proofErr elements are now preserved)
	runCount := 0
	var firstRun *Run
	for _, c := range np.Children {
		if r, ok := c.(*Run); ok {
			runCount++
			if firstRun == nil {
				firstRun = r
			}
		}
	}
	if runCount != 1 {
		t.Fatal("expected only one run but has", runCount)
	}
	if len(firstRun.Children) != 1 {
		t.Fatal("expected only one run.child but has", len(firstRun.Children))
	}
	if firstRun.Children[0].(*Text).Text != allmergedtext {
		t.Fatal("expected merged text [", allmergedtext, "] but has [", firstRun.Children[0].(*Text).Text, "]")
	}
	np = p.MergeText(MergeSamePropRuns)
	// Count only Run children
	runCount = 0
	for _, c := range np.Children {
		if _, ok := c.(*Run); ok {
			runCount++
		}
	}
	if runCount != 13 {
		t.Fatal("expected 13 runs but has", runCount)
	}
	sb := strings.Builder{}
	for _, c := range np.Children {
		r, ok := c.(*Run)
		if !ok {
			continue // skip non-Run elements like ProofErr
		}
		if len(r.Children) > 1 {
			t.Fatal("expected 0/1 run.child but has", len(r.Children))
		}
		if len(r.Children) == 1 {
			sb.WriteString(r.Children[0].(*Text).Text)
		}
		sb.WriteString("|")
	}
	if sb.String() != propmergedtext {
		t.Fatal("expected merged text [", propmergedtext, "] but has [", sb.String(), "]")
	}
	np = p.MergeText(MergeSamePropRunsOf("Bold", "Size", "Underline"))
	// Count only Run children
	runCount = 0
	for _, c := range np.Children {
		if _, ok := c.(*Run); ok {
			runCount++
		}
	}
	if runCount != 7 {
		t.Fatal("expected 7 runs but has", runCount)
	}
	sb.Reset()
	for _, c := range np.Children {
		r, ok := c.(*Run)
		if !ok {
			continue // skip non-Run elements like ProofErr
		}
		if len(r.Children) > 1 {
			t.Fatal("expected 0/1 run.child but has", len(r.Children))
		}
		if len(r.Children) == 1 {
			sb.WriteString(r.Children[0].(*Text).Text)
		}
		sb.WriteString("|")
	}
	if sb.String() != namedpropmergdtext {
		t.Fatal("expected merged text [", namedpropmergdtext, "] but has [", sb.String(), "]")
	}
}
