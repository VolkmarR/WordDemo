import './Home.css'

type HomeProps = {
    file: File | null
    onFileSelected: (file: File | null) => void
    onOpenPreview: () => void
}

export default function Home({
                                 file,
                                 onFileSelected,
                                 onOpenPreview,
                             }: HomeProps) {
    return (
        <div className="home">
            <div className="home-card">
                <h1>Word Preview POC</h1>

                <p>
                    Select a DOCX document and then compare the different rendering
                    approaches.
                </p>

                <input
                    type="file"
                    accept=".docx"
                    onChange={(e) =>
                        onFileSelected(e.target.files?.[0] ?? null)
                    }
                />

                {file && (
                    <div className="selected-file">
                        <h3>Selected file</h3>

                        <div>Name: {file.name}</div>
                        <div>Size: {Math.round(file.size / 1024)} KB</div>

                        <button
                            type="button"
                            className="open-button"
                            onClick={onOpenPreview}
                        >
                            Open Preview
                        </button>
                    </div>
                )}
            </div>
        </div>
    )
}