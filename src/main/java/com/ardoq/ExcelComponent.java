package com.ardoq;

import com.ardoq.model.Component;

import java.util.ArrayList;
import java.util.HashMap;

/**
 * Created by magnulf on 01.07.14.
 */
public class ExcelComponent {

    static HashMap<String, ExcelComponent> pathMap = new HashMap<String, ExcelComponent>();

    public static ArrayList<ExcelComponent> getRootNodes() {
        return rootNodes;
    }

    static ArrayList<ExcelComponent> rootNodes = new ArrayList<ExcelComponent>();
    private final String path;

    public ExcelComponent(String path, Component myComponent, ExcelComponent parent) {
        this.myComponent = myComponent;
        this.path = path;
        pathMap.put(path, this);
        if (parent != null) {
            this.parent = parent;
            this.parent.addChild(this);
        } else {
            rootNodes.add(this);
        }
    }

    public ExcelComponent(String path, Component myComponent) {
        this(path, myComponent, null);
    }

    public ArrayList<ExcelComponent> getChildren() {
        return children;
    }

    public void addChild(ExcelComponent myComponent) {
        this.children.add(myComponent);
    }

    public Component getMyComponent() {
        return myComponent;
    }

    public void setMyComponent(Component myComponent) {
        this.myComponent = myComponent;
    }

    public String getPath() {
        return path;
    }

    public static ExcelComponent getExcelComponent(String path) {
        return pathMap.get(path);
    }


    Component myComponent;
    ExcelComponent parent;
    ArrayList<ExcelComponent> children = new ArrayList<ExcelComponent>();


}
