(function(window, undefined) {
    // 初始化插件
    window.Asc.plugin.init = function() {
        // 处理文档内容已准备就绪事件
        function handleContentReady() {
            // 获取文档内容
            var documentContent = Api.GetDocumentContent();
            console.log("文档内容："+documentContent)
            // 提取新增文本
            var newText = extractNewText(documentContent);

            // 实体识别和提取
            var entities = extractEntities(newText);

            // 高亮文本实体并在实体下方展示框
            highlightEntitiesAndShowContent(entities);
        }

        // 注册文档内容已准备就绪事件监听器
        this.attachEvent("onDocumentContentReady", handleContentReady);
    };

    window.Asc.plugin.button = function(id) {
        // 处理按钮点击事件（如果有需要）
    };

    // 提取新增文本的函数
    function extractNewText(documentContent) {
        // TODO: 根据实际需求，实现从文档内容中提取新增文本的逻辑
        // 您可以将当前文档内容与之前存储的内容进行比较，以确定新增的文本部分

        // 为了演示目的，我们假设整个文档内容都被视为新增文本
        return documentContent;
    }

    // 提取实体的函数
    function extractEntities(text) {
        var entities = [];
        var regex = /公司/g; // 使用正则表达式匹配"公司"
        var match;

        while ((match = regex.exec(text)) !== null) {
            var entity = {
                startIndex: match.index, // 实体的起始位置
                endIndex: match.index + match[0].length - 1 // 实体的结束位置
            };
            entities.push(entity);
        }

        return entities;
    }

    // 高亮文本实体并在实体下方展示框
    function highlightEntitiesAndShowContent(entities) {
        var oDocument = Api.GetDocument();

        // 遍历所有实体
        for (var i = 0; i < entities.length; i++) {
            var entity = entities[i];

            // 创建一个文本范围对象来标记实体位置
            var range = oDocument.CreateRange(entity.startIndex, entity.endIndex);

            // 设置高亮样式
            range.SetHighlightColor("#00bbff"); // 设置为黄色高亮

            // 在实体下方创建一个框
            var oFrame = Api.CreateFrame();
            oFrame.SetWidth("100px"); // 设置框的宽度
            oFrame.SetHeight("50px"); // 设置框的高度
            oFrame.SetPosition(entity.startIndex, entity.endIndex + 1); // 设置框的位置，+1 是为了放在实体下方

            // 创建框内的内容
            var oParagraph = Api.CreateParagraph();
            oParagraph.AddText("山东龙羲含章");
            oFrame.InsertContent([oParagraph]);

            // 将范围和框插入文档
            oDocument.InsertContent([range, oFrame]);
        }
    }
})(window, undefined);