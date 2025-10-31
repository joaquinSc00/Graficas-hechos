from flask import Flask, request, jsonify

app = Flask(__name__)


def clamp(value, min_value, max_value):
    return max(min_value, min(max_value, value))


@app.post("/layout")
def layout():
    data = request.get_json(force=True)
    limits = data.get("style_limits", {})
    body_limits = limits.get("body", {})
    title_limits = limits.get("title", {})

    body_base = body_limits.get("base", 9.5)
    body_min = body_limits.get("min", body_base)
    body_max = body_limits.get("max", body_base)
    title_base = title_limits.get("base", 25.0)
    title_min = title_limits.get("min", title_base)
    title_max = title_limits.get("max", title_base)

    notes = data.get("notes", [])
    response = {"instructions": []}

    if not isinstance(notes, list):
        return jsonify(response)

    for note in notes:
        note_id = note.get("id") or f"nota{len(response['instructions'])+1}"
        chars = int(note.get("chars", 0))
        has_photo = bool(note.get("hasPhoto", False))

        body_pt = body_base
        title_pt = title_base
        action = "keep"

        unplaced_flag = bool(note.get("unplaced", False))
        has_frame = note.get("hasFrame", True)

        if unplaced_flag or has_frame is False:
            action = "unplaced"
        else:
            if chars > 1800:
                body_pt = clamp(body_base - 0.5, body_min, body_max)
                action = "expand_frame_first"
            elif 1400 <= chars <= 1800:
                body_pt = clamp(body_base - 0.25, body_min, body_max)
                action = "tighten"
            elif chars < 900 and not has_photo:
                body_pt = clamp(body_base + 0.25, body_min, body_max)
                action = "loosen"

        body_pt = clamp(body_pt, body_min, body_max)
        title_pt = clamp(title_pt, title_min, title_max)

        response["instructions"].append({
            "id": note_id,
            "body_pt": body_pt,
            "title_pt": title_pt,
            "action": action
        })

    return jsonify(response)


if __name__ == "__main__":
    app.run(port=5001)
