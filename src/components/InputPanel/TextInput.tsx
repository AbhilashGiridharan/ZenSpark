interface Props {
  value: string;
  onChange: (v: string) => void;
  placeholder?: string;
  label?: string;
  maxLength?: number;
}

const SOFT_LIMIT = 20_000;

export default function TextInput({
  value,
  onChange,
  placeholder = "Paste notes, RFP content, requirements, meeting minutes…",
  label,
  maxLength = 40_000,
}: Props) {
  const len = value.length;
  const isOver = len > SOFT_LIMIT;

  return (
    <div>
      {label && (
        <label className="mb-1 block text-xs font-medium text-gray-500">
          {label}
        </label>
      )}
      <textarea
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        maxLength={maxLength}
        rows={5}
        className="w-full resize-none rounded-lg border border-gray-300 bg-gray-100/40 px-3 py-2 text-sm text-gray-800 placeholder-gray-400 focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600"
        spellCheck={false}
      />
      <div className={`mt-0.5 text-right text-xs ${isOver ? "text-amber-400" : "text-gray-400"}`}>
        {len.toLocaleString()} chars
        {isOver && ` · Large context — only the first ${SOFT_LIMIT.toLocaleString()} chars will be sent`}
      </div>
    </div>
  );
}
