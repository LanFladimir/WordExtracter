package bean;

/**
 * 选择的文件
 * WORD/PDF
 *
 */
public class SelectFile {
    private String name;
    private String path;

    public SelectFile(String name, String path) {
        this.path = path;
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPath() {
        return path;
    }

    public void setPath(String path) {
        this.path = path;
    }
}
