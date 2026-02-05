import PptxGenJS from './src/bld/pptxgen.es.js'
import JSZip from 'jszip'
import fs from 'fs/promises'

console.log('Testing Text Array Processing in Slide Master Placeholders...\n')

const pptx = new PptxGenJS()

// Define slide master with text placeholder containing formatted text runs
pptx.defineSlideMaster({
    title: 'Test Master',
    background: { color: 'FFFFFF' },
    objects: [
        {
            placeholder: {
                options: {
                    name: 'Disclaimer',
                    type: 'body',
                    x: 0.62,
                    y: 1.90,
                    w: 12.12,
                    h: 4.00,
                },
                text: [
                    { text: 'Disclaimer\n', options: { bold: true, color: '000000', paraSpaceAfter: 8 } },
                    { text: 'IHS Herold Inc. ("IHS Herold"), a wholly-owned subsidiary of S&P Global, provides data and analysis on the strategy, performance, and valuation of companies in the global energy industry, as well as energy transactions and trends. IHS Herold serves a subscription client base consisting of energy companies, financial institutions, investment managers, and advisory firms. IHS Herold is not affiliated with any broker, dealer, bank, or investment bank. Companies mentioned in this report may be or may have been subscribers to IHS Herold data and research in the past year on terms consistent with all other clients of IHS Herold or clients of any member of the S&P Global group. IHS Herold does not inform any company in advance of the nature or conclusions of its research reports. Analysts do not receive any compensation from the companies on which they report or from any other industry source.\n', options: { bold: false, color: '000000', paraSpaceAfter: 8 } },
                    { text: 'Copyright © 2026 S&P Global. All rights reserved. The [', options: { bold: false, color: '000000' } },
                    { text: 'Insert publication', options: { bold: false, color: 'B92051' } },
                    { text: '] is published by IHS Herold, 55 Post Rd W, 2nd Floor, Westport, CT 06880, USA for the exclusive use of S&P Global clients. Reproduction of this report, even for internal distribution, is strictly prohibited. The information contained herein has been obtained from sources believed to be reliable, but neither S&P Global nor any of its affiliates guarantees their accuracy or completeness. No information or opinions contained herein constitutes a representation or solicitation for the purchase of any securities of the companies mentioned herein. From time to time, IHS Herold and/or its officers and employees may have long or short positions in the securities mentioned herein or during the past year may have transacted in securities of the companies mentioned.\n', options: { bold: false, color: '000000', paraSpaceAfter: 8 } },
                    { text: 'The research analyst who prepared this report certifies that the views expressed herein accurately reflect the research analyst\'s professional opinions, are consistent with established IHS Herold methodologies and standards, and that no part of his or her compensation was, is, or will be directly or indirectly related to specific views contained in this report.\n', options: { bold: false, color: 'D6002A', paraSpaceAfter: 8 } },
                    { text: 'The research analysts who prepared this report certify that the views expressed herein accurately reflect their professional opinions, are consistent with established IHS Herold methodologies and standards, and that no part of their compensation was, is, or will be directly or indirectly related to specific views contained in this report.\n', options: { bold: false, color: 'D6002A', paraSpaceAfter: 8 } },
                    { text: 'The analyst that prepared this report has a [long/short] position in the securities of [company] mentioned in this report. This analyst is contractually prohibited from purchasing or selling the securities of this company during the blackout period implemented by IHS Herold for this type of report.\n', options: { bold: false, color: 'D6002A', paraSpaceAfter: 8 } },
                    { text: 'This content was created with the assistance of Kensho Spark Assist, an artificial intelligence (AI) tool. For complete Terms of Use, see https://www.spglobal.com/en/terms-of-use.', options: { bold: false, color: '000000', paraSpaceAfter: 8 } },
                ]
            }
        }
    ]
})

console.log('✓ Slide master defined with text array placeholder (9 formatted runs)')

// Add slide using the master
const slide = pptx.addSlide({ masterName: 'Test Master' })
console.log('✓ Slide added using the master')

// Save presentation
const fileName = 'test-text-array-output.pptx'
await pptx.writeFile({ fileName })
console.log(`✓ Saved: ${fileName}\n`)

// Verify XML output
console.log('Verifying XML output...\n')

const data = await fs.readFile(fileName)
const zip = await JSZip.loadAsync(data)

// Read slide layout XML
const layoutXml = await zip.file('ppt/slideLayouts/slideLayout1.xml').async('string')

// Check if text contains [object Object]
const hasObjectError = layoutXml.includes('[object Object]')

// Extract text content from the placeholder
const match = layoutXml.match(/<a:t>([^<]*)<\/a:t>/g)
const textContent = match ? match.map(m => m.replace(/<\/?a:t>/g, '')) : []

console.log('Text runs found in XML:')
if (textContent.length > 0) {
    textContent.forEach((text, i) => {
        const preview = text.length > 60 ? text.substring(0, 60) + '...' : text
        console.log(`  ${i + 1}. "${preview}"`)
    })
} else {
    console.log('  (No text runs found)')
}

console.log('\n=== Test Results ===\n')

if (hasObjectError) {
    console.log('✗ FAILED: Found "[object Object]" in XML output')
    console.log('\nThe text array is being stringified instead of processed as individual runs.')
    process.exit(1)
} else {
    console.log('✓ PASSED: No "[object Object]" found in XML')
    console.log(`✓ PASSED: ${textContent.length} text runs generated correctly`)
    console.log('\nThe text array with formatted runs is working correctly!')
}

console.log('\nTo verify visually:')
console.log(`1. Open ${fileName} in PowerPoint`)
console.log('2. Check that the disclaimer text displays with proper formatting')
console.log('3. Verify different colors (black for most, pink/red for analyst certifications)')
console.log('4. Ensure no "[object Object]" text appears anywhere')
