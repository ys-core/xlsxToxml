# Git常用命令

////This part content comes from Master branch

1. get reset --soft HEAD^ 可撤销最近那一次到本地仓库的commit，不删除工作空间改动代码，撤销commit，不撤销git add . 与get reset --soft HEAD~1命令等效 

---

2. get reset --soft HEAD~2 可撤销最近两次到本地仓库的commit

---

3. get reset --mixed HEAD^ 不删除工作空间改动代码，撤销commit，并且撤销git add . 操作，与默认的 get reset HEAD^ 命令等效

---

4. get reset --hard HEAD^ 删除工作空间改动代码，撤销commit，撤销git add . 完成这个操作后，就恢复到了上一次的commit状态

---

5. 顺便说一下，如果commit注释写错了，只是想改一下注释，只需要 git commit --amend 此时会进入默认vim编辑器，修改注释完毕后保存即可

---

6. 要覆盖远端的版本信息，使远端的仓库版本也回退到相应的版本可以加上参数 --force
   git push origin master --force**
   
---
7. 