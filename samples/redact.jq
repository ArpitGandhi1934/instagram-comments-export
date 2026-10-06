# Applied to the raw Actor output before committing:
#   jq -f samples/redact.jq out/raw.json > samples/nasa-reel-top-comments.json
#
# The sample is real Actor output with every commenter pseudonymized:
# - Each commenter's username becomes commenter_01, commenter_02 ... (in order of first
#   appearance), verified or not. Their user id and avatar URL are set to null.
# - Comment ids become comment_01, comment_02 ...
# - Comment text is cut to its first sentence (plus "…" when something was cut). @mentions become
#   @user, commenters' usernames anywhere in the text become their alias, emails are removed.
# The post (code, URL, owner account) and every count, date and flag are left untouched.

def rx: gsub("(?<c>[.*+?^${}()|\\[\\]\\\\/])"; "\\\(.c)");

def firstSentence:
  (gsub("\\s+"; " ") | sub("^ "; "") | sub(" $"; "")) as $t
  | ([$t | capture("^(?<s>.+?[.!?。！？])( |$)")] | first | .s?) as $s
  | if $s != null and ($s | length) <= 300 then ($s + (if ($s | length) < ($t | length) then " …" else "" end))
    elif ($t | split(" ") | length) > 25 then ($t | split(" ")[:25] | join(" ")) + " …"
    else $t end;

([.[] | select(.dataType == "comment") | .owner.username] | reduce .[] as $u ([]; if index([$u]) then . else . + [$u] end)) as $users
| def alias($u): "commenter_" + ((($users | index([$u])) + 1) | tostring | if length < 2 then "0" + . else . end);
  def pad($n): ($n | tostring | if length < 2 then "0" + . else . end);
  [ to_entries[] | .key as $i | .value
    | if .dataType == "comment" then
        .owner.username as $u
        | .id = "comment_" + pad($i + 1)
        | .owner.username = alias($u)
        | .owner.id = null
        | .owner.profilePicUrl = null
        | .text |= (
            if type == "string" then
              (reduce $users[] as $x (.; if ($x | length) > 3
                  then gsub("(?<![A-Za-z0-9_.])" + ($x | rx) + "(?![A-Za-z0-9_])"; alias($x)) else . end))
              | gsub("[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}"; "(email removed)")
              | gsub("@[A-Za-z0-9_.]+"; "@user")
              | firstSentence
            else . end)
      else . end ]
