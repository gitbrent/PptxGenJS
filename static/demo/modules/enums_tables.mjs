export const TABLE_AUTO_PAGE_TEST = `
-- Slide 1: Introduction and Overview --
--------------------------------------------------
1. Welcome to the auto-paging table reproduction test.
2. This file is meant to trigger the "needs repair" bug.
3.    * Note: Ensure trailing spaces are preserved.
4.  - Line with a dash bullet and extra spaces at end.
5. Here is a sample line with a tab indent:		Indented info.
6. Data point: Value1   (with extra spaces)
7. Data point:	Value2	  (with tab indentation)
8. Multiple spaces are included in this line:      many spaces.
9. Here’s a note: Remember to check all bullet levels.
10.	• A bullet using a tab and dot at the beginning.
11. Extra line with a mix of tabs	 and multiple spaces.
12. End-of-line spaces are critical here.
13. Line with mixed formatting:	- Indented dash bullet.
14. A very detailed point:
	    * Sub-point with tab indent and extra space at the end.
15. This line ends with spaces and a tab:
16. A simple sentence with irregular spacing.
17. Testing line-break handling:
This is a new line immediately after.
18. Lines that use "quotes" can be tricky.
19. Trailing spaces:
20.   * Another bullet with excessive indent and spacing.
21. Some lines are intentionally long to test auto-paging. This is line number 21.
22. Mixing dashes and bullets:   - A bullet with a dash and a tab:		Extra detail.
23. Another sample with extra formatting.
24. A line with a spaced bullet:   •   Note with extra spaces.
25. 	Item with leading tab before text.
26. Text with multiple    spaces between words.
27. More content to simulate real table data.
28. A line with a mix of symbols:  *-+*   random symbols.
29. Repeat a similar structure: 		Indented again.
30. Tab-started line with trailing spaces:	    Ending in spaces.
31. This is the 31st line with enough content to test auto-paging.
32. More and more text, building up the slide content.
33. Example:    A bullet with extra indent.
34. Testing:    "Quotes and spaces" with trailing spaces.
35. Mixing tabs and spaces in one line:	Spaces then tab.
36. Multiple consecutive spaces are visible here:           many.
37. Note: blank lines may appear after this sentence.
38. Here is a bullet:  - Simple dash bullet.
39. An indented bullet with a tab:	• Tabbed bullet example.
40. More detailed line:    Step 1: initialize, Step 2: process, Step 3: finish.
41. A line that uses both spaces and a tab:   	        Mixed indent.
42. This line is for testing wrapping: It should be long enough to force an auto-break in some table viewers.
43. Additional detail:    Check if line-breaks are preserved correctly.
44. A well-formed line with bullet:   • Final note for this section.
45. Sequence number line: 45 - with a dash and extra spacing.
46. Reminder: Always check for trailing spaces on each line.
47. More content to simulate text density in a slide.
48. A decorative bullet:     • — decorative dash and bullet mix.
49. Space-filled line:
50. End of Slide 1 content.

-- Slide 2: Detailed Bullet and Indent Tests --
--------------------------------------------------
1. Beginning Slide 2: Focus on detailed bullet formatting.
2. This slide contains bullets, indents and various formatting challenges.
3. 	- Item with a leading tab and dash bullet.
4.	• Item with a tab and bullet point.
5.         * Item with multiple spaces before an asterisk.
6. Data point: Value A with trailing spaces.
7. Data point:	Value B with tab spacing.
8. Extra spaces inserted between words:    For example, this text.
9. Quote "Tested line" with trailing spaces.
10. A line intentionally split after a colon:
    Details continue on the next line.
11. Bulleted list start:
		• Bullet sub-item A.
12. Continuation of bullet:   - Bullet sub-item B.
13. Mixing indentation:
        * This line starts with 8 spaces.
14. Tabs before dash:	- Tabbed dash item.
15. Lines with multiple indented bullet styles:
	- Primary bullet.
		* Secondary bullet with extra indent.
16. Repeat bullet styles:
	• First style again.
17. Consistency test:  Data line with  varying spacing.
18. Adding a note:  // Ensure formatting remains intact.
19. A line ending with a tab and spaces:
20. Extra content: More words to push the line length.
21. A dash bullet with additional content:	- Dash bullet and long text for auto-paging.
22. More bullet tests:        * Indented star bullet.
23. Another bullet line with mix:	- Mixed bullet using a tab.
24. Notice this line uses a colon followed by a space: Data continues.
25. A well-spaced bullet:    • Clear bullet point.
26. Check the combination of spaces and tabs:		Confusing format?
27. An indented note:	    "Remember: spacing matters!"
28. More data: Value C, Value D, with proper formatting.
29. Data sequence: 1, 2, 3, 4, 5 - with irregular spacing.
30. Mixing different bullet characters:	•, -, * all in one.
31. A line of text followed by a blank line.
32. A line with a trailing period and extra spaces.
33. Data row example: A1, B1, C1 with trailing spaces.
34. A tabbed bullet:  	- Tab-indented list item.
35. Indentation test:	          This line begins with many spaces.
36. A comment-like line: // This line is a comment.
37. A symbolic line: @#$%^&*() symbols in text.
38. The auto-paging bug should be reproducible with this long text.
39. Indented bullet and text:		• Make sure to check all indents.
40. This is a line with trailing tab and space:
41. A subtle mix:     "Quotes" with some spaces.
42. A line meant for testing line breaks: First part,
    Second part continues on a new line.
43. Another bullet item follows:
	• With tab indent and extra spacing.
44. A line with an explicit indent marker: >> Indented information.
45. Extra note: Some lines have irregular and extra tab spaces.
46. Bullets nested:
	- Level 1
		- Level 2 with more details.
47. More detailed bulleting: * Nested bullet style.
48. Data summary: Key figures: 123, 456, 789, with spaces.
49. Final testing line in Slide 2 with lots of emphasis.
50. End of Slide 2 content.

-- Slide 3: Mixed Content and Spacing --
--------------------------------------------------
1. Start Slide 3: Combining paragraphs and bullet points.
2. This slide integrates multiple formatting elements, including line breaks.
3. Paragraph begins here: The purpose of this slide is to mix formatting styles.
4. Bulleted list:
	- First bullet item in Slide 3.
5. Continue with a bullet:	* Second bullet, with a tab indent.
6. A paragraph with extra line breaks follows:
   This is the first line of a multi-line paragraph.
7. Second line of the paragraph: Notice the spacing.
8. Third line:    Additional indented text for emphasis.
9. End of paragraph.
10. Empty line follows:
11. Another paragraph: Setting up for further tests.
12. Here is a dash bullet:    - Testing bullet consistency.
13. A star bullet with extra spaces:         * Emphasized bullet text.
14. Mixed bullet with colon:	• Item with colon: details follow.
15. Tab and space mix line:	    Data point with tab indent.
16. A long line with multiple spacing issues to check auto-wrapping functionality in tables.
17. Consider this a random note: Tabs, spaces, and extra indentations.
18. Line with a trailing space: That ends with spaces.
19. A note with special characters: !@#$%^&*() included.
20. Quote test: "This is a quoted statement" with trailing spaces.
21. Another bullet:   - Mixed formatting bullet.
22. A line with both space and tab at the beginning:	     Indented message.
23. Numbered list start:
    1. First numbered item.
24.    2. Second numbered item with extra indentation.
25.    3. Third numbered item, which might cause auto-break.
26. A descriptive line: Text continues on the next line.
27. Second part of the description: properly indented.
28. An indented quote:
	    "A famous quote with extra tabs at the start."
29. Another line, using dash and bullet:    - Intersection of symbols: * and •.
30. Line showing random punctuation: ...???!!!
31. Example line with significant trailing spaces:
32. Another line starting with a tab:	Starts with a tab.
33. Additional bullet:  • Follow-up bullet point.
34. Testing a mix of indentation and line breaks:
    First half with text, then a break.
35. Second half on new line, but retains indentation.
36. More content for testing.
37. A long line that might require wrapping: This is a long narrative line intended to test whether auto-paging and text wrapping behave as expected when the content overflows the table's boundaries.  A long line that might require wrapping: This is a long narrative line intended to test whether auto-paging and text wrapping behave as expected when the content overflows the table's boundaries. A long line that might require wrapping: This is a long narrative line intended to test whether auto-paging and text wrapping behave as expected when the content overflows the table's boundaries.
38. Extra indentation again:	    Level two indentation test.
39. Spacer line: (just spaces)
40. More bullet details:   * Important: Check every detail.
41. Note on formatting: Mixing tabs, spaces, and multiple line breaks.
42. A combination line: Start bullet - then colon: followed by details.
43. More detailed instructions: Indent all following lines.
44. A placeholder line: Lorem ipsum dolor sit amet, consectetur adipiscing elit.
45. Continued placeholder line: Vestibulum vitae.
46. More extra space:  	* With a combination of symbols.
47. Another line: Ensure each bullet's spacing is varied.
48. Line with mix:	  - Test bullet with multiple spaces.
49. Final note: Check the file for auto-repair message triggers.
50. End of Slide 3 content.

-- Slide 4: Tabbed and Indented Lines --
--------------------------------------------------\\\\
1. Begin Slide 4 focusing on tabs and indents.
2. This slide's primary purpose is to test heavy tabbing.
3.	• Tabbed bullet item at the start.
4.	- A dash bullet with a leading tab.
5.	    * A star bullet indented with multiple tabs.
6. A line starting with several tabs:		    Data begins here.
7. More content with extra spaces after a tab indent.
8. A line that mixes tabs and spaces:		Data arranged with both.
9. Tabbed numerical list:
	1. Item one with a tab.
10.	2. Item two under a numerical list with a tab.
11. Another line with tabs:		This line continues the tab pattern.
12. A double-tabbed line:		    Indented even more.
13. Testing line: Here is a bullet with both tabs and spaces at the start.
14. A line ending with both tabs and spaces:
15. Yet another tab-focused line:		Tabs, then extra spaces.
16. More indent tests: Line with heavy tab indent at beginning.
17. Line with tab indents and multiple space separation between words.
18. Data example with mixed formatting:		Data point followed by details.
19. This is just another tab-introduced line for the test.
20. Insert a line with only tabs:
21. Insert a line with only spaces:
22. A combined line:	    Mixed content after tab and space.
23. A line with a tab followed by a dash bullet:	- Detail following a tab.
24. Another bullet with tab indent:		• Another detail.
25. A line with a repeated tab sequence:
			    Repeated indentation test.
26. More tabs:		- Dash bullet with excessive tab spacing.
27. Data row with tabs and spaces:		A  B  C  D entries.
28. A line with extra tabs:			Ending in tab space.
29. Testing a very long tabbed sentence that should be broken into multiple segments if the auto-paging logic is applied correctly in the application.
30. A line mixing bullets, tabs, and spaces:	    	• Mixed and complicated.
31. More tab indents:		This line checks strict tab behavior.
32. Another bullet after a tab:		- Continuing bullet list.
33. Testing a blank tabbed line:
34. Observing how multiple tabs act before text:		Tabs in front of text here.
35. A line with a tab, then bullet, then text:		• Bullet then text with tab.
36. Another format:		- Dash bullet with text following a tab.
37. Data point:	Information provided with tab indentation.
38. A line with interleaved tabs and spaces:	    Data with a pattern.
39. Testing sequential tabs:		\t\tSimulated tab output.
40. More content with a heavy tab start:		    Leading tab message.
41. A comment-styled line with tab indent:	// Tabbed comment.
42. Another comment:		// Followed by a second tabbed comment.
43. A line combining tabs and quotes:		"Tabbed quote test."
44. More indented content:		    Lorem ipsum dolor sit amet.
45. A triple tabbed bullet:				* Deeply indented bullet.
46. A heavily indented line:				    Multiple tabs followed by text.
47. A mixed bullet:		• Bullet prefixed with multiple tabs.
48. Testing a line with tabs at the beginning and extra trailing spaces.
49. Another test line:	    Tabs and spaces combine for testing.
50. End of Slide 4 content.

-- Slide 5: Multi-Line Content and Spacing --
--------------------------------------------------
1. Slide 5 starts with multi-line paragraph testing.
2. This paragraph is split across several lines to emulate a long block of text.
3. First line of the paragraph: The quick brown fox jumps over the lazy dog.
4. Continued thought: Pack my box with five dozen liquor jugs.
5. Further explanation: How razorback-jumping frogs can level six piqued gymnasts.
6. A line with a bullet inside the paragraph:	• Inserted bullet detail.
7. The paragraph continues with more spaced text and multiple indents.
8. Here is another thought, split over several parts:
	This is the second segment of the paragraph, following a break.
9. Third segment:		Notice the indentation for subsequent lines.
10. A concluding line of the multi-line paragraph with trailing spaces.
11. A new paragraph starts here with additional details.
12. Bullet within the paragraph:
    - Detail point within a running paragraph.
13. More continuation: The text must preserve line breaks.
14. A line with a tab indent and extra spacing:
		    Annotated line with multiple indent levels.
15. A detailed explanation line with a mix of spaces:
    Explanation continues with a colon: Details follow.
16. More text to simulate content:    Additional data, more text, and extra spaces.
17. Random content: Lorem ipsum dolor sit amet, consectetur adipiscing elit.
18. Continuation: Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.
19. A line with both bullet and numbers: 1. First item followed by 2. Second item.
20. Text with trailing spaces to test the auto-paging artifact.
21. More multi-line testing: This is an extra long sentence designed to exceed typical line lengths and force wrapping in various table renderers.
22. Another paragraph with irregular line breaks: Start here then
continue without proper punctuation.
23. A line with a dash bullet embedded:    - Embedded dash bullet.
24. Continue further: The quick brown fox jumps.
25. Insert a break:

26. A line with spaces and a tab mixed:	    	Testing indent after a blank line.
27. More points:
	    * Star bullet after a blank.
28. Repeating format:    ... and then more text to ensure the line is long enough.
29. A line with a colon and extra spacing:    Detail: followed by extended description.
30. A data row with multiple columns: A, B, C, D, E.
31. Further text: Remember to include various bullet formats.
32. Another tab-indented note:		Tabbed note for replication test.
33. Continuation of paragraph: More content, more text, and more spacing.
34. Another bullet:   • This bullet is for additional testing.
35. Adding extra spaces:    Check line endings carefully.
36. More mixed content: This line tests auto-page logic with lengthy text.
37. An indented line with mixed symbols:		- Symbols: @!#$%^&*()
38. A quoted line: "Testing line with quotes" with trailing spaces.
39. Extra spaces and line breaks are essential here.
40. A numbered list within multi-line paragraph:
    1. First point continues.
41.    2. Second point with more text.
42. More detailed line:    Indentation and bullets show up correctly.
43. A line with multiple spaces between words: Word    spacing    test.
44. More text to fill the slide with varied content.
45. Next line testing auto-break compatibility in rendering systems.
46. Line with just a tab and space:
47. A placeholder line: Continue simulating multiple line breaks.
48. More bullets:      - Another bullet test.
49. Yet another detailed line: Insert content and ensure trailing spaces are present.
50. End of Slide 5 content.

-- Slide 6: Summary and Recap --
--------------------------------------------------
1. Final Slide: Summarizing all the points tested.
2. This slide revisits bullet points, indentation, and line breaks.
3. A recap bullet:   • The auto-paging bug may be triggered by mixed spacing.
4. Another recap point:  - Consistent use of tabs and spaces.
5. Text review: Each slide tests a variety of formatting elements.
6. Numbered list recap:
    1. Introduction was filled with mixed bullets.
7.    2. Detailed bullet tests checked indentation.
8.    3. Mixed content verified line breaks.
9.    4. Heavy tabs in Slide 4 were crucial.
10. Recap continuing with further points.
11. A review bullet:   * Ensure trailing spaces are always present.
12. Additional reminder: Verify if auto-paging fails on text with extra formatting.
13. Reiterated note: Tabs, spaces, and line breaks are all significant.
14. Extra line with a focus on punctuation:    Check commas, colons, and dashes.
15. A short line just to show a break:
16. Another tabbed bullet recap:	- Review bullet with tab indent.
17. Summary continuation: Final confirmation of all formatting issues.
18. A line with mixed formatting:	• Mixed bullet and dash.
19. Re-emphasize: The reproduction test is complete.
20. End of numbered list in this slide.
21. More recap: Review the auto-paging sections carefully.
22. Notice the detailed bullet:   * Final bullet for checking.
23. A line with both dashes and spaces:    - A conclusive dash.
24. Final commentary: Formatting inconsistencies could lead to repair warnings.
25. Assurance line: Each line is designed to test rendering behavior.
26. Multiple tabbed lines now:		• Another final bullet point.
27. More spacing:     Final emphasis on extra spaces and trailing spaces.
28. Another recap step: Detailed formatting is critical.
29. A tabbed note:		This note is indented for style.
30. Final review: Ensure that every bullet, line break, tab, and space is accounted for.
31. Concluding note:   The test file has reached a critical mass.
32. More detailed bullet:   - Extra detailed point for review.
33. Combining text and symbols: Check for formatting errors.
34. Additional recap bullet:    • Ensure auto-paging logic kicks in.
35. Tabbed format recap:	- Confirm that lines wrap as expected.
36. A long recap line: This sentence is designed to trigger auto-breaks in tables when the content overflows the set boundaries.
37. Another mix:    "Quotes" and symbols are part of the test.
38. A concluding bullet with space:   * Final check on trailing spaces.
39. Counting line: Yet another line to ensure hundreds of lines are achieved.
40. More data: Testing auto-paging functionality continues.
41. An indented summary:	    With multiple levels.
42. Another final note:    Each slide's end is marked clearly.
43. Penultimate line: A close look at detailed testing conditions.
44. Extra spacing check:      Some lines end with many spaces.
45. One more line:   Testing a large block of text for reproducibility.
46. Final bullet confirmation:	• This is the final bullet.
47. Last call:  The reproduction file is nearly complete.
48. Concluding summary: A final emphasis on mixed formatting styles.
49. Final verification: Check for table auto-paging repair triggers.
50. End of Slide 6 content and end of reproduction text file.
 `;
