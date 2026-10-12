# {{ Report.Title }}

{{ Report.Author }} — {{ Report.Date }}

{{ Report.Summary }}

## Findings
{% for finding in Report.Findings %}
- **{{ finding.Area }}** ({{ finding.Status }}) — {{ finding.Owner }}
{% endfor %}

## Actions
{% for action in Report.Actions %}
### {{ action.Title }}

{{ action.Detail }}
{% endfor %}
