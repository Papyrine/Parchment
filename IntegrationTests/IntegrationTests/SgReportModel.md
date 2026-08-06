# Report for {{ Customer.Name }}

{% for line in Lines %}
- {{ line.Description }}
{% endfor %}
