import os
import xmind
import pandas
import shutil
import rarfile
import zipfile


def extract(d_path, f_path, mode="zip"):
    """
    zip解压缩乱码问题处理
    :param d_path:
    :param f_path:
    :return:
    """
    root = d_path
    if not os.path.exists(root):
        os.makedirs(root)

    if mode == 'zip':
        zf = zipfile.ZipFile(f_path, "r")
    elif mode == 'rar':
        zf = rarfile.RarFile(f_path, "r")
    else:
        raise Exception("mode error")

    for n in zf.infolist():
        srcName = n.filename
        try:
            decodeName = srcName.encode("cp437").decode("utf-8")
        except:
            try:
                decodeName = srcName.encode("cp437").decode("gbk")
            except:
                decodeName = srcName
        spiltArr = decodeName.split("/")
        path = root
        for temp in spiltArr:
            path = os.path.join(path, temp)

        if decodeName.endswith("/"):
            if not os.path.exists(path):
                os.makedirs(path)
        else:
            if not os.path.exists(os.path.dirname(path)):
                os.makedirs(os.path.dirname(path))
            f = open(path, "wb")
            f.write(zf.read(srcName))
            f.close()
    zf.close()


def aftertreatment(path):
    """
    **場景　xmind8 可以打开　xmind2020 报错
    main_fest.xml(xmind8 打开另存后 更改后缀为.zip  里边包含META-INF/manifest.xml)
    xmind 修改后缀为zip ----》解压---- 》放入main_fest.xml  --- 》压缩zip  修改后缀为xmind**
    """
    # 修改名字
    retval = os.path.dirname(os.path.abspath(__file__))
    folder = os.path.dirname(path)
    name = os.path.basename(path)
    unzip_folder = os.path.splitext(name)[0]
    zip_name = unzip_folder + ".zip"
    os.chdir(folder)
    os.rename(name, zip_name)
    os.chdir(retval)
    # 解压
    unzip_path = str(os.path.join(folder, unzip_folder))
    if not os.path.exists(unzip_path):
        os.makedirs(unzip_path, exist_ok=True)

    inf_folder = os.path.join(unzip_path, "META-INF")
    if not os.path.exists(inf_folder):
        os.mkdir(inf_folder)

    extract(unzip_path, os.path.join(folder, zip_name))
    shutil.copyfile("./META-INF/manifest.xml", os.path.join(inf_folder, "manifest.xml"))
    os.remove(os.path.join(folder, zip_name))
    shutil.make_archive(unzip_path, 'zip', unzip_path)
    file_path = unzip_path + '.zip'
    os.chdir(os.path.dirname(file_path))
    os.rename(os.path.basename(file_path), name)
    os.chdir(retval)
    shutil.rmtree(unzip_path)


def excel_to_xmind(excel_path, xmind_path):
    # 读取Excel文件
    df = pandas.read_excel(excel_path, sheet_name="用例")
    workbook = xmind.load("./templ/testcase.xmind")
    sheet = workbook.getPrimarySheet()
    root_topic = sheet.getRootTopic()
    root_topic.setTitle(df["所属产品"].values[0])
    model_topics = {}
    # 遍历Excel中的每一行
    for _, row in df.iterrows():
        # 创建用例分支
        if row['所属模块'] not in model_topics:
            model_topic = root_topic.addSubTopic()
            model_topic.setTitle(row['所属模块'])
            model_topics[row['所属模块']] = model_topic
        else:
            model_topic = model_topics[row['所属模块']]
        case_topic = model_topic.addSubTopic()
        case_topic.setTitle(row['用例标题'])
        case_topic.addMarker(f"priority-{row['优先级']}")

        # 添加基本属性
        props = {
            "相关需求": row["相关需求"],
            "分支": row["分支"],
            "用例类型": row["用例类型"],
            "适用阶段": row["适用阶段"],
            "前置条件": row["前置条件"] if pandas.notna(row["前置条件"]) else "无",
        }
        for k, v in props.items():
            attr_topic = case_topic.addSubTopic()
            attr_topic.setTitle(k)
            attr_topic.addSubTopic().setTitle(v)

        # 添加步骤与预期结果
        steps_topic = case_topic.addSubTopic()
        steps_topic.setTitle("步骤和预期")
        steps = [s.split(".", 1)[-1] for s in str(row["步骤"]).split("\n") if s.strip()]
        expects = [s.split(".", 1)[-1] for s in str(row["预期"]).split("\n") if s.strip()]
        for i, (step, expect) in enumerate(zip(steps, expects), 1):
            step_topic = steps_topic.addSubTopic()
            step_topic.setTitle(f"步骤{i}: {step}")
            step_topic.addSubTopic().setTitle(f"预期结果: {expect}")

    # 保存XMind文件
    xmind.save(workbook, xmind_path)
    # 修复xmind文件
    aftertreatment(xmind_path)
