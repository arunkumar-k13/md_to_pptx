# Markdown to PPTX Conversion Test

## 1. Basic Text Formatting

This is a normal paragraph containing standard text to establish a baseline for font size and style.

This is bold text used to emphasize importance.

This is italic text used for emphasis or titles.

This is bold and italic text combining both styles.

~~This is strikethrough text~~ indicating deleted or obsolete information.

Inline code is used for variable names like myVariable or System.out.println().

Text containing special characters: ! @ # $ % ^ & * ( ) _ + = { } [ ] < > ? / \ | ~

Text containing numbers, dates, percentages, currency values, and decimal values:
* Numbers: 1234567890
* Dates: 2024-05-20, 12/31/1999
* Percentages: 45.5%, 100%, 0.001%
* Currency: $19.99, £50.00, €100.50, ¥5000, ₹1,200.00
* Decimals: 3.14159265, 0.00001, -45.2

Unicode characters: © ® ™ € £ ¥ ₹ ₩ ± × ÷ ≠ ≤ ≥ → ← ↑ ↓ ★ ✓ ✗

Emoji: 😀 🚀 🔥 ✅ ⚠️ ❌ 💡 📊 🎯

## 2. Headings

# Heading Level 1
This is the primary title of a document or slide.

## Heading Level 2
This represents a major section or a new slide title.

### Heading Level 3
This is a subsection within a major section.

#### Heading Level 4
This is a sub-subsection for more granular detail.

##### Heading Level 5
This is a deep sub-section, often used for specific notes.

###### Heading Level 6
This is the smallest heading level available in Markdown.

## 3. Paragraphs

Short text paragraph.

Medium text paragraph: This paragraph contains several sentences to test how the layout engine handles standard blocks of text within a slide. It should maintain proper line spacing and margin alignment.

Long text paragraph: This paragraph is designed to test the limits of text containment. It contains a very long sentence to test how the conversion pipeline handles text wrapping and potential overflow when a single line of logic extends beyond the typical boundaries of a standard PowerPoint text box container without any manual line breaks being provided by the user.

Multiple punctuation marks: Is this working? Yes! (Mostly). "Check this out," he said; it's quite impressive: 1. Start, 2. End. [Brackets], {Braces}, and <Greater/Less>.

## 4. Lists

### Unordered List

* First item
* Second item
* Third item
    * Nested item
    * Another nested item
        * Deeply nested item
        * Another deeply nested item

### Ordered List

1. First step
2. Second step
3. Third step
    1. Nested step
    2. Another nested step
        1. Deeply nested step
        2. Another deeply nested step

### Mixed List

1. Main item
    * Sub item
    * Another sub item
2. Second main item
    * Sub item
        1. Nested numbered item
        2. Another numbered item

## 5. Blockquotes

> This is a basic blockquote.

> This is a longer blockquote containing multiple sentences.
> It should test how quoted content is handled during PPTX conversion.

> Important: This blockquote also contains bold text and inline code.

## 6. Links

* [OpenAI](https://www.openai.com)
* [GitHub](https://github.com)
* [Example Website](https://example.com)
* https://www.example.com
* [Link with Query](https://example.com/search?q=markdown+test&lang=en#results)

## 7. Code Blocks

### Java

```java
public class MarkdownTest {
    public static void main(String[] args) {
        System.out.println("Hello from Markdown!");
    }
}
```

### JavaScript

```javascript
function calculateTotal(items) {
    return items.reduce((total, item) => total + item.price, 0);
}
```

### Python

```python
def calculate_average(numbers):
    return sum(numbers) / len(numbers)
```

### SQL

```sql
SELECT id, user_name, email
FROM users
WHERE user_name LIKE '%test%';
```

### JSON

```json
{
  "id": 1001,
  "name": "Markdown Test",
  "active": true,
  "values": [10, 20, 30]
}
```

## 8. Tables

| ID | Name | Status | Score |
| -: | ------- | --------- | ----: |
| 1 | Alice | Active | 95 |
| 2 | Bob | Pending | 82 |
| 3 | Charlie | Completed | 91 |
| 4 | David | Failed | 47 |

| Feature | Description | Status | Notes |
| ---------- | --------------------------------- | --------- | ----------------------------------- |
| API | Retrieves Markdown content | Completed | Endpoint working correctly |
| Parser | Parses Markdown | Completed | Multiple Markdown constructs tested |
| PPTX | Converts Markdown into PowerPoint | Testing | Formatting verification required |
| Validation | Checks generated presentation | Pending | Requires manual verification |

## 9. Complex Table

| Metric | Value | Percentage | Cost | Date | Result |
| ------- | -----: | ---------: | -----------: | ---------- | --------- |
| Users | 12,450 | 87.5% | ₹1,25,000.50 | 2026-08-18 | ✓ Pass |
| Errors | 125 | 1.2% | $1,250.75 | 2026-08-17 | ⚠ Warning |
| Success | 12,325 | 98.8% | €9,999.99 | 2026-08-16 | ✓ Pass |

## 10. Horizontal Rules

---

Then continue with another section.

---

Then continue with more content.

## 11. Images

![Sample Image](https://via.placeholder.com/800x400.png)

![Another Image](https://via.placeholder.com/600x300.png)

## 12. Nested Formatting

Bold with `inline code`

Italic with bold text

Bold italic with `inline code`

> Bold quote with italic text and inline code

## 13. Mathematical and Technical Content

E = mc²

a² + b² = c²

x = (-b ± √(b² - 4ac)) / 2a

O(n)

O(log n)

O(n²)

HTTP 200 OK

HTTP 404 Not Found

HTTP 500 Internal Server Error

## 14. Special Characters and Unicode

ABCDEFGHIJKLMNOPQRSTUVWXYZ

abcdefghijklmnopqrstuvwxyz

0123456789

! @ # $ % ^ & * ( ) - _ + =

{ } [ ] < > / \ | : ; " ' , . ? ~

© ® ™ ° ± × ÷ ≠ ≤ ≥ ∞ √ ∑ ∆ µ Ω

₹ $ € £ ¥ ₩

→ ← ↑ ↓ ↔ ⇒ ⇐

✓ ✔ ✗ ✘ ★ ☆

😀 😃 😄 😁 😎 🚀 🔥 💡 🎯 📊 ⚠️ ✅ ❌

## 15. Long Content

Project "Aetheris" is a fictional cloud-native infrastructure management suite designed to automate the deployment of multi-region Kubernetes clusters. This section serves as a stress test for content distribution.

The first phase of the project involves the development of the "Core-Nexus" module, which handles state synchronization across geographically dispersed nodes. This ensures that even in the event of a total regional outage, the metadata remains consistent and available for recovery operations in surviving regions.

The second phase introduces "Sentinel-Flow," an AI-driven monitoring engine that predicts potential bottlenecks by analyzing historical traffic patterns. By utilizing machine learning algorithms, the system can proactively scale resources before the actual demand spikes occur, thereby maintaining high availability and low latency for the end-users.

Finally, the project concludes with the "Aetheris-Portal," a unified dashboard providing real-time visibility into the health of the entire infrastructure. This dashboard integrates metrics, logs, and traces into a single pane of glass, allowing site reliability engineers to diagnose issues rapidly and perform root cause analysis without switching between multiple disconnected tools.

* This list item is here to test list rendering inside a long content section.
* It checks if the conversion tool properly calculates slide height and wraps the list to a new slide if necessary.
* Maintaining the heading placement and paragraph spacing during an automatic slide split is a critical success factor for this test.

## 16. Short Content

# Short Section
Hello.

## Another Short Section
Test.

## 17. Empty-Like Content

## 18. Realistic Project Report

### Executive Summary
Project "Synth-Link" is a bridge between legacy mainframe data and modern cloud-based analytics platforms.

### Objectives
* Migrate 500TB of historical data.
* Ensure zero downtime during the cutover phase.
* Implement end-to-end AES-256 encryption.
* Reduce data retrieval latency by 40%.
* Provide a RESTful API for third-party integrations.

### Architecture
The system follows a microservices architecture deployed on AWS EKS, utilizing an Event-Driven design pattern via Apache Kafka for data ingestion.

### Technology Stack
* Java
* Spring Boot
* React
* MySQL
* MongoDB
* REST API
* Docker

### API Flow
1. Client sends request.
2. Backend validates request.
3. Service retrieves Markdown.
4. Markdown is processed.
5. PPTX is generated.
6. PPTX is returned to the client.

### Test Results

| Test Case ID | Description | Input | Expected Result | Actual Result | Status |
| :--- | :--- | :--- | :--- | :--- | :--- |
| TC-01 | Auth Check | Valid Token | 200 OK | 200 OK | Pass |
| TC-02 | Data Pull | ID 101 | JSON Object | JSON Object | Pass |
| TC-03 | MD Parse | Bold Text | Text | Text | Pass |
| TC-04 | PPTX Gen | MD Content | .pptx file | .pptx file | Pass |
| TC-05 | Unicode | Emoji 🚀 | Rendered | Rendered | Pass |
| TC-06 | Table Render | 4x4 Table | Correct Grid | Correct Grid | Pass |
| TC-07 | Image Embed | URL Link | Image Shown | Image Shown | Pass |
| TC-08 | Code Block | Java Code | Highlighted | Highlighted | Pass |
| TC-09 | Nested List | 3 Levels | Indented | Indented | Pass |
| TC-10 | Link Click | URL | Open Browser | Open Browser | Pass |

### Risks
* Data Corruption: Mitigated by checksum validation during transfer.
* Latency Spikes: Mitigated by implementing Redis caching.

### Conclusion
The project successfully demonstrates the feasibility of high-speed Markdown to PPTX conversion.

## 19. Edge Cases

# This is an extremely long heading that is designed to see if it wraps or if it extends past the boundaries of the PowerPoint slide which would be quite problematic for the layout

This is an extremely long paragraph without any breaks intended to see if the conversion tool handles word-wrapping correctly or if the text just disappears into the void of the slide's right margin.

| Very Long Table Cell Header |
| :--- |
| This cell contains an exceptionally long string of text to test how the table cell width and height adjust dynamically to the content provided within the Markdown source. |

* This is a very long list item that contains enough text to span across multiple lines and potentially push subsequent list items onto a new slide entirely if not handled with care by the conversion logic.

# Heading 1
## Heading 2
### Heading 3

0, -1, 1.23456789, 999999999999
100%, 99.99%, 0.01%
₹0.00, ₹1,25,000.50, $999.99, €1,234.56

## 20. Final Validation Section

# Conversion Test Complete

The Markdown document should be converted into a PPTX while preserving:
1. Heading hierarchy
2. Paragraphs
3. Bold formatting
4. Italic formatting
5. Strikethrough formatting
6. Inline code
7. Lists
8. Nested lists
9. Blockquotes
10. Links
11. Code blocks
12. Tables
13. Images
14. Unicode characters
15. Special characters
16. Mathematical expressions
17. Long content
18. Short content
19. Slide structure
20. Overall readability
